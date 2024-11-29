using Microsoft.Office.Interop.Access.Dao;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System.Runtime.CompilerServices;
using Word = Microsoft.Office.Interop.Word;
// Imports Word library.

/*
 * The idea is to be able to input a list
 * containing each color, what should go where that color is,
 * and the document that should be edited.
 */

//Potential way store and read data used in the creation of 
// template documents.
Dictionary<string, string> inputData = new();



/**
 * Temporary method used for testing.
 */
void TESTINGInitializer()
{
    //Input how the data may be stored.
    inputData.Add("DEFENDANT_NAME", "John Doe");
    inputData.Add("DOB", "11/24/1995");
    inputData.Add("COUNT_NUMBER", "1");
    inputData.Add("OFFENSE_DATE", "3/8/2024");
    inputData.Add("CONDUCT", "Bad");
    inputData.Add("VICTIM", "Ronald Roe");

    //Print out the data stored in the dictionary.
    foreach (KeyValuePair<string, string> entry in inputData){
        Console.WriteLine($"{entry.Key}: {entry.Value}");
    }
}

TESTINGInitializer();







//C:\Users\shuff\source\repos\AutoDocsCSharp\AutoDocsDraft\bin\Debug\net8.0\testdocuments\CSharpDocTest.docx
//Placeholder directory used, this will be changed later to a permanent address.
string directory = "C:\\users\\shuff\\documents\\csharpdocs\\CSharpDocTest.docx";
string directory2 = "C:\\users\\shuff\\documents\\csharpdocs\\Assault.docx";

//Variable used for the creation of a new Word application so that we can use methods on it.
var wordApp = new Word.Application();

    //Shows the document when editing for debugging purposes, will be False later.
    // be sure to add .close if set to false.
wordApp.Visible = true;

    //Adds a new document to the Word application.
var docx = wordApp.Documents.Open(directory);

    //Creates the selection of the document as a variable.
var selection = wordApp.Selection;


Word.Range docRange = docx.Content;






























/* LIST EVERY BOLDENED WORD
docRange.Find.ClearFormatting();
docRange.Find.Forward = true;
docRange.Find.Format = true;
docRange.Find.Font.Bold = 1;
docRange.Find.Execute(FindText: "");

while(docRange.Find.Found){
    string boldText = docRange.Text;
    Console.WriteLine($"-- {boldText.TrimEnd()}");
    docRange.Find.Execute(FindText: "");
}
*/






/* FIND A BOLD WORD AND UNBOLDEN IT
selection.Find.ClearFormatting();
selection.Find.Font.Bold = 1;
if (selection.Find.Execute(FindText: "", Format: true)){
    string foundText = selection.Text;
    Console.WriteLine(foundText);
    selection.Font.Bold = 0;
}
*/






/**
 * A function that we can use to find and replace any 
 *  instance of a word with whatever we desire.
 */
void findReplaceText(){

        //This foreach loops allows us to go through every 
        // key value pair, with the key being the word to replace
        // and the value being the replacement word.
    foreach (KeyValuePair<string,string> entry in inputData) {
        
            //We clear the formatting here to prevent any 
            // unintended issues (not sure what it does).
        selection.Find.ClearFormatting();
        selection.Find.Replacement.ClearFormatting();
        
            //We set the statement to replace all instances of a word
            // as a simple replaceAll object.
        object replaceAll = Word.WdReplace.wdReplaceAll;

            //This statement allows us to find every instance of
            // the entry's Key and replace it with the entry's Value.
        selection.Find.Execute(FindText: entry.Key, ReplaceWith: entry.Value, Replace: replaceAll);

            //Debug statement to make sure that the foreach loop is ran.
        Console.WriteLine("Replaced!");
    }

        //This statement lets us set the date in the document, it's not included in the
        // for loop because it will not be a part of input data.
    selection.Find.Execute(FindText: "TODAYS_DATE", ReplaceWith: DateTime.Now.ToString("d"));
}

//findReplaceText();






/* PRINTS OUT EVERY COLOR AVAILABLE

for (int colRange = 1; colRange < 19; colRange++)
{
    selection.Font.ColorIndex = (WdColorIndex)colRange;

    selection.TypeText($"{selection.Font.Color}\n");
}

*/



/* CODE USED TO FIND A STRING AND REPLACE IT IN BOLD "you've been found"
object findText = "replace me!";
selection.Find.ClearFormatting();
selection.Font.Name = "Verdana";
selection.Font.Size = 12;
selection.Font.Bold = 1;
if (selection.Find.Execute(findText))
{
    Console.WriteLine("FOUND IT!");
    selection.TypeText("you've been found \n");
    selection.Find.ClearFormatting();
    selection.Find.Text = "replace me!";
    selection.Find.Replacement.ClearFormatting();
    selection.Find.Replacement.Text = "replaced!";
    object replaceAll = Word.WdReplace.wdReplaceAll;
    selection.Find.Execute(FindText: "replace me!", ReplaceWith: "replaced!", Replace: replaceAll);
}
else {
    Console.WriteLine("DIDN'T FIND IT!");
}



*/



/*
    //Writes given text into the document.
selection.TypeText("Hello to all!");

    //Saves the document at a specified directory.
docx.SaveAs(directory);
*/
