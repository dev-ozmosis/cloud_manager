const should = require("should");
const MSDrive = require("./../connections/ms_drive").MSDrive;
const { ms_drive_config } = require("./../settings");

describe("Testing MSDrive", () => {

    it("Should try to connect to MSConnection", (done) => {

        const drive = new MSDrive();
        drive.connect()
            .then ((success) => {
                should.equal(success, true);
                should.notEqual(MSDrive.session.authStr, "");
                done();
            })
            .catch ((err)=> { done(); });
    });

    it("Should try to get MSUserInfo", (done) => {

        const drive = new MSDrive();
        drive.getUserInfo(ms_drive_config.UserPrincipalName)
            .then ((success) => {
                done();
            })
            .catch ((err)=> { done(); });
    });

    it("Should try to get MSDriveRootDirectory", (done) => {

        const drive = new MSDrive();
        drive.getRoot()
            .then ((success) => {
                done();
            })
            .catch ((err)=> { done(); });
    });

    it("Should try to get MSChildsFromDirectory", (done) => {

        const drive = new MSDrive();
        drive.getRoot()
            .then ((rootFiles) => {

                if (rootFiles.length === 0) { 
                    return [];
                }

                const foundDirectories = rootFiles.filter( (item) => { return item.isFolder === true} );
                should.notEqual(foundDirectories, null);
                should.notEqual(foundDirectories.length, 0);

                return drive.itemsFromDirectory(foundDirectories[0].id);
            })
            .then ((directoryItems) => {

                if (directoryItems.length !== 0) { 
                    
                    should.notEqual(directoryItems, null);
                    should.notEqual(directoryItems.length, 0);
                }

                done();
            })
            .catch ((err)=> { done(); });
    });

    it("Should try to create MSCreateFolder", (done) => {

        const directoryName = "test3";
        const drive = new MSDrive();
        drive.getRoot()
            .then ((rootFiles) => {

                if (rootFiles.length === 0) { 
                    return;
                }

                const foundDirectories = rootFiles.filter( (item) => { return item.isFolder === true} );
                should.notEqual(foundDirectories, null);
                should.notEqual(foundDirectories.length, 0);

                return drive.createFolder(directoryName, foundDirectories[0].id);
            })
            .then ((directory) => {

                should.notEqual(directory, null);
                should.equal(directory.name, directoryName);

                done();
            })
            .catch ((err)=> { done(); });
    });
});