const qs = require("qs"); 
const axios = require("axios");
const moment = require("moment");
const { ms_drive_config } = require("./../settings");
const { UserInfo, MSDriveItem } = require("./ms_drive_models");

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

    getUserInfo = async (userEmail) => {

        try {

            await this.connect();

            const url = "https://graph.microsoft.com/v1.0/users/" + userEmail;
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

            const url = "https://graph.microsoft.com/v1.0/users/"+ this.user.id + "/drive/root/children";  
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

    itemsFromDirectory = async (directoryId) => {

        try {

            await this.connect();

            const url = "https://graph.microsoft.com/v1.0/users/"+ this.user.id + "/drive/items/" + directoryId + "/children";  
            const response = await axios.default.get(url, { headers: { Authorization: MSDrive.session.authStr } });
            console.log("[MSDrive::itemsFromDirectory]: ", response.data);
        }
        catch (err) {
            
            const response = err.response;
            if (response.data.error != undefined && response.data.error != null) {

                const error = response.data.error;
                console.log("[MSDrive::itemsFromDirectory]: ", error);
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