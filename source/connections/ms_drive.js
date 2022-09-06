import { ms_drive_config } from "./../settings";

class MSDrive {

    connect = async () => { 

        if (MSDrive.session.authStr === "") {

            const now = moment();
            const startToken = MSDrive.session.startIn;
            const endToken = moment(startToken).add(MSDrive.session.expiresIn, 's'); 
            if (now < endToken) {
                return;
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
        }
        catch (err) {

            console.log("[MSDrive::connect] Error: ", err);
        }
    }
}

MSDrive.session = {
    authStr: "",
    expiresIn: 0,
    startIn: null
};

export default MSDrive;