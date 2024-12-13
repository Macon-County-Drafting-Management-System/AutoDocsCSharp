

    //This variable keeps our json contents for use within the website.
const docJson = {
    PropertiesList: {Defendant: "John Doe", DOB: "12/5/2004", 
        "Count Number": "8", "Offense Date": "12/7/2024", 
        Conduct: "No good", Victim: "Ronald Roe"}
};

    //This variable is the download button used in the website.
const downloadButton = document.querySelector("button[class='download']");


    //This block of variables represents every input from the html
    // and are named accordingly to the type of input.
const defendantName = document.querySelector("input[id='Defendant']");
const defendantDOB = document.querySelector("input[id='DOB']");
const countNumber = document.querySelector("input[id='Count Number']");
const offenseDate = document.querySelector("input[id='Offense Date']");
const conduct = document.querySelector("input[id='Conduct']");
const victim = document.querySelector("input[id='Victim']");

    //This function is used to apply our inputs to the json object
    // that was created earlier.
function applyInputs(){
        
        //Each variable in this block corresponds to a json 
        // value.
    docJson.PropertiesList["Defendant"] = defendantName.value;
    docJson.PropertiesList["DOB"] = defendantDOB.value;
    docJson.PropertiesList["Count Number"] = countNumber.value;
    docJson.PropertiesList["Offense Date"] = offenseDate.value;
    docJson.PropertiesList["Conduct"] = conduct.value;
    docJson.PropertiesList["Victim"] = victim.value;
}

    //This function is used to take our json string and
    // download it to the user's computer in json format.
function downloadJSON(){

        //We stringify the json object so that it can be pasted into
        // the json document.
    const jsonString = JSON.stringify(docJson);
    
        //We take our json string and place it in a Binary Large Object with the 
        // json file extension and create it into a url link.
    const blob = new Blob([jsonString], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    
        //We create an arbitrary element "a".
    const a = document.createElement("a");

        //We set this element's hyperlink reference to the blob
        // url, then proceed to download the object it as a 
        // data.json file and click.
    a.href = url;
    a.download = "data.json";
    a.click();
}

    //When the download button is clicked, we both apply all 
    // inputs to the json file and proceed to download it.
downloadButton.onclick = () => {
    applyInputs();
    downloadJSON();
}

testButton.onclick = () => {
    applyInputs();
}