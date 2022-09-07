const { MSDrive } = require("./connections/ms_drive");

function getRootFiles() {
    const drive = new MSDrive();
    drive.getRoot()
        .then( (rootItems) => {
            
            const ulElement = document.getElementById("listItems");
            for (let index = 0; index < rootItems.length; index++)
            {
                let opt = rootItems[index];
                const li = document.createElement("li");
                li.innerHTML = opt.name
               
                ulElement.appendChild(li);
            }	
        })  
        .catch ((err) => {
            console.log ("err: " + JSON.stringify(err));
        });
}

const rootBtn = document.getElementById("rootBtn");
rootBtn.addEventListener("click", (e) => {

    getRootFiles();
});

