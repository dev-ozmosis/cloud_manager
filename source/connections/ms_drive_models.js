class UserInfo {

    constructor(id, name) {
        this.id = id;
        this.name = name;
    }
}

class MSDriveItem {

    constructor(id, name, createdAt) {

        this.id = id;
        this.name = name;
        this.createdAt = createdAt;
        this.createdBy = null;
        this.parentId = "";
        this.isFolder = false;
    }

    static createFromJSON(json) {

        let item = new MSDriveItem(json.id, json.name, json.createdAt);
        if (json.createdBy !== undefined && json.createdBy !== null) {

            const user = json.createdBy.user;
            item.createdBy = new UserInfo(user.id, user.displayName);
        }

        if (json.parentReference !== undefined && json.parentReference !== null) {
            item.parentId = json.parentReference.id;
        }

        item.isFolder = (json.folder !== undefined && json.forder !== null);

        return item;
    }
}

module.exports.MSDriveItem = MSDriveItem;
module.exports.UserInfo = UserInfo;