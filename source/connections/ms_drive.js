const qs = require("qs"); 
const axios = require("axios");
const moment = require("moment");
const { ms_drive_config } = require("./../settings");
const { UserInfo, MSDriveItem } = require("./ms_drive_models");
const { URLS } = require("./../constants");

//https://docs.microsoft.com/en-us/graph/api/resources/driveitem?view=graph-rest-1.0
class MSDrive {

    constructor() {
        this.user = null;
    }

    connect = async () => { 

        if (MSDrive.session.authStr !== "") {

            const now = moment();
            const startToken = MSDrive.session.startIn;
            const endToken = moment(startToken).add(MSDrive.session.expiresIn, 'seconds'); 
            if (now < endToken) {
                return true;
            }
        }

        axios.defaults.headers.post['Content-Type'] = 'application/x-www-form-urlencoded';
        const postData = {
            client_id: ms_drive_config.appId,
            scope: ms_drive_config.MS_GRAPH_SCOPE,
            client_secret: ms_drive_config.appSecret,
            grant_type : "client_credentials"
        };

        try {

            const response = await axios.default.post(ms_drive_config.tokenEndpoint, qs.stringify(postData));

            //FIXME: is the best option store the data in this way?
            MSDrive.session.authStr = response.data.token_type  + " " + response.data.access_token;
            MSDrive.session.expiresIn = response.data.expires_in;
            MSDrive.session.startIn = moment();

            //TODO: read the email from file settings or from user interaction
            this.user = await this.getUserInfo(ms_drive_config.UserPrincipalName); 
        }
        catch (err) {

            const response = err.response;
            if (response.data.error != undefined && response.data.error != null) {

                const error = response.data.error;
                console.log("[MSDrive::connect] error: ", error);
            }
            return false
        }

        return true;
    }

    getUserInfo = async (userEmail = "") => {

        if (userEmail === "") {
            return null;
        }
        
        try {

            await this.connect();

            const url = URLS.MS_BASE_URL + userEmail;
            const response = await axios.default.get(url, { headers: { Authorization: MSDrive.session.authStr } });

            const data = response.data;
            let user = new UserInfo(data.id, data.displayName);
            
            console.log("[MSDrive::getUserInfo] user: ", user);
            return user;
        }
        catch (err) {

            const response = err.response;
            if (response.data.error != undefined && response.data.error != null) {

                const error = response.data.error;
                console.log("[MSDrive::getUserInfo] err: ", error);
            }

            return null;
        }
    }

    getRoot = async () => {

        try {

            await this.connect();

            const url = URLS.MS_BASE_URL + this.user.id + "/drive/root/children";  
            const response = await axios.default.get(url, { headers: { Authorization: MSDrive.session.authStr } });

            let items = [];
            const data = response.data;
            for (let i = 0; i < data.value.length; i++) {

                const item = MSDriveItem.createFromJSON(data.value[i]);
                items.push(item);
            }

            return items;
        }
        catch (err) {
            
            const response = err.response;
            if (response.data.error != undefined && response.data.error != null) {

                const error = response.data.error;
                console.log("getRoot error: ", error);
            }
        }
    }

    itemsFromDirectory = async (directoryId = "") => {

        if (directoryId === "") { 
            return; 
        }

        try {

            await this.connect();

            const url = URLS.MS_BASE_URL + this.user.id + "/drive/items/" + directoryId + "/children";  
            const response = await axios.default.get(url, { headers: { Authorization: MSDrive.session.authStr } });

            let items = [];
            const data = response.data;
            for (let i = 0; i < data.value.length; i++) {

                const item = MSDriveItem.createFromJSON(data.value[i]);
                items.push(item);
            }

            console.log("[MSDrive::itemsFromDirectory]: ", JSON.stringify(items));
            return items;
        }
        catch (err) {
            
            const response = err.response;
            if (response.data.error != undefined && response.data.error != null) {

                const error = response.data.error;
                console.log("[MSDrive::itemsFromDirectory]: ", error);
            }
        }
    }

    
    createFolder= async (name, parentId = "") => {

        //https://docs.microsoft.com/en-us/graph/api/driveitem-post-children?view=graph-rest-1.0&tabs=javascript
        if (name === "") { return null; }

        try {

            await this.connect();

            let url = URLS.MS_BASE_URL + this.user.id + "/drive/items/" + parentId;
            if (parentId === "") {
                //TODO: if s empty can be root directory, get root directory id!
            }

            url += "/children";

            const directoryParam = { "name": name, "folder": { } };
            axios.defaults.headers.post['Content-Type'] = 'application/x-www-form-urlencoded';
            const response = await axios.default.post(url, directoryParam, { headers: { Authorization: MSDrive.session.authStr } });

            const data = response.data;
            const folderItem = MSDriveItem.createFromJSON(data);
            
            return folderItem;
        }
        catch (err) {

            const response = err.response;
            if (response.data.error != undefined && response.data.error != null) {

                const error = response.data.error;
                console.log("[MSDrive::createFolder]: ", error);
            }
        }
    }
}

MSDrive.session = {
    authStr: "",
    expiresIn: 0,
    startIn: null
};

module.exports.MSDrive = MSDrive;