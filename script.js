const msalConfig = {
    auth: {
        clientId: "8942b408-2646-4bbd-bb4f-7a82226351c2",
        authority: "https://login.microsoftonline.com/9eaa6d01-27d2-4363-9466-ebe0ef39b381",
        redirectUri: "https://precisiondisciplinesystem.github.io/discipline-management-system/"
    }
};


const msalInstance = new msal.PublicClientApplication(msalConfig);


async function signIn() {

    try {

        const loginResponse = await msalInstance.loginPopup({

            scopes: [
                "User.Read"
            ]

        });


        console.log("Signed in user:");

        console.log(loginResponse.account);


        alert(
            "Welcome " + loginResponse.account.name
        );


    }

    catch(error){

        console.error(error);

        alert(
            "Login failed. Check console for details."
        );

    }

}
function loadGenerator(){


    if(!document.getElementById("employeeSelect")){
        return;
    }


    employees.forEach(employee=>{

        let option =
        document.createElement("option");

        option.value = employee.employeeNumber;

        option.text =
        employee.name +
        " - " +
        employee.employeeNumber;


        employeeSelect.appendChild(option);

    });



    articles.forEach(article=>{


        let option =
        document.createElement("option");


        option.value =
        article.code;


        option.text =
        article.code +
        " - " +
        article.description;


        articleSelect.appendChild(option);


    });



    loadProgressions();

}



function loadProgressions(){


    progressionSelect.innerHTML="";


   let article =
articles.find(
    a=>a.code === document.getElementById("articleSelect").value
);


if(!article){
    return;
}


    article.progressions.forEach(p=>{


        let option =
        document.createElement("option");


        option.text=p;


        progressionSelect.appendChild(option);


    });


}



function generateDocument(){


    let date =
    new Date()
    .toISOString()
    .slice(0,10)
    .replaceAll("-","");


    let article =
    articleSelect.value;


    let progression =
    progressionSelect.value
    .toUpperCase();


    let filename =
    date +
    "-" +
    article +
    " " +
    progression;


    document.getElementById("result")
    .innerText =
    "Document Created: "
    + filename;


}



window.onload=function(){

    loadGenerator();

};

if(document.getElementById("articleSelect")){

    document.getElementById("articleSelect")
    .addEventListener(
        "change",
        loadProgressions
    );

}
