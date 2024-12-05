using Microsoft.Office.Interop.Access.Dao;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System.IO;
using System.Runtime.CompilerServices;
using Word = Microsoft.Office.Interop.Word;
using System.Text.Json;
// Imports Word library.

namespace AutoDocsDraft {

    public class DocumentProperties{
        public Dictionary<string, string>? PropertiesList { get; set; } = [];
    }


    public class Program{


        
        public static void Main() {


                //Potential way store and read data used in the creation of 
                // template documents.
            Dictionary<string, List<string>>? docPropertiesCC = [];

            



                //C:\Users\shuff\source\repos\AutoDocsCSharp\AutoDocsDraft\bin\Debug\net8.0\testdocuments\CSharpDocTest.docx
                //Placeholder directories used, this will be changed later to a permanent address.
            string directory = "C:\\users\\shuff\\documents\\csharpdocs\\CSharpDocTest.docx";
            string directory2 = "C:\\users\\shuff\\documents\\csharpdocs\\Assault.docx";
            string directory3 = "C:\\users\\shuff\\documents\\csharpdocs\\AssaultCC.docx";

                //Variable used for the creation of a new Word application so that we can use methods on it.
            Word.Application wordApp = new() {
                    //Shows the document when editing for debugging purposes, will be False later.
                    // be sure to add .close if set to false.
                Visible = true
            };

            
             //Adds a new document to the Word application.
            var docx = wordApp.Documents.Open(directory3);

                //Creates the selection of the document as a variable.
            var selection = wordApp.Selection;

                //Creates the document range as a variable so that
                // calling is made quicker.
            Word.Range docRange = docx.Content;



            /**
             * This function reads a template for it's Content Controls
             *  and sets the name of the content control, it's index, and 
             *  placeholder value in a dictionary for use later.
             * 
             * Dictionary form- {"Content Control Name": ["Content Control Index", "Placeholder Value"]}
             */
            void readTemplate() {


                //For loop that allows us to iterate through every content
                // control inside the document.
                for (int i = 1; i <= docRange.ContentControls.Count; i++) {

                    //We set the content control's range as a variable
                    // for easily editing.
                    Word.Range CCselection = docRange.ContentControls[i].Range;

                    //Here, we add the text of the selection, it's placeholder value,
                    // as the key for our dictionary, with it's CC index and placeholder
                    // text string as the two entries in the list.
                    docPropertiesCC.Add(CCselection.Text, [$"{i}", $"Placeholder {i}"]);
                }

                /* DEBUG FUNCTION
                    //This foreach loop lists out every key and their values
                    // for us. 
                foreach (KeyValuePair<string, List<string>> entry in docPropertiesCC){
                    Console.WriteLine($"{entry.Key}:{entry.Value[0]}, {entry.Value[1]}");
                }
                */
            }


            /**
             * This function allows us to populate a word document based off of our 
             *  list of content controls, and what text should replace it.
             */
            void populateDocument() {


                    //This foreach loop allows us to go through our entire list of input data 
                    // and populate the document based off of our scheme used.
                foreach (KeyValuePair<string, List<string>> entry in docPropertiesCC) {

                        //This if statement singles out the entry containing Today's Date, since that 
                        // is a special value that changes each day and thus has it's own variable.
                        //It could change though.
                    if (entry.Key.Contains("Today")) {

                            //Sets the Today's Date content control as today's date.
                        docRange.ContentControls[int.Parse(entry.Value[0])].Range.Text = DateTime.Now.ToString("d");
                    } else {

                            //Sets the corresponding content control to it's new value.
                        docRange.ContentControls[int.Parse(entry.Value[0])].Range.Text = entry.Value[1];
                    }
                }
            }




            /**
             * This method will read a json file and serialize the data for us 
             * into a usable class with a dictionary.
             */
            void readJSON(bool readFile = true, bool updateCCProperties = true) {

                if (readFile) {
                    try {
                        string jsonFilePath = "C:\\users\\shuff\\OneDrive\\Documents\\jsonTest.json";


                        using StreamReader sr = new StreamReader(jsonFilePath);
                        string jsonString = sr.ReadToEnd();
                        DocumentProperties docProp = JsonSerializer.Deserialize<DocumentProperties>(jsonString);
                        Console.WriteLine(docProp.PropertiesList.Count());


                        foreach (KeyValuePair<string,string> property in docProp.PropertiesList) {
                            if (docPropertiesCC.ContainsKey(property.Key)) {
                                docPropertiesCC[property.Key].Insert(1, property.Value);
                            }
                        }




                    } catch (Exception e) {
                        Console.WriteLine($"Oops!:\n{e.ToString()}");
                    }

                    
                } else {

                    string jsonString = 
                        """      
                        {
                          "PropertiesList": {
                            "Defendant": "John Doe",
                            "DOB": "12/5/2004",
                            "Count Number": "8",
                            "Offense Date": "12/7/2024",
                            "Conduct": "No good",
                            "Victim": "Ronald Roe"
                          }
                        }
                        """;
                    
                    DocumentProperties? docProp = JsonSerializer.Deserialize<DocumentProperties>(jsonString);
                    
                }




            }





            readTemplate();
            readJSON();
            populateDocument();












            //wordApp.Quit();
        }

        


        /**
         * These are methods that I am keeping as a reference but 
         *  not currently using to achieve the endgoal of drafting 
         *  a document.
         */
        void OverflowMethodologies() {

            
            Dictionary<string, string>? inputData = [];



            /**
             * Temporary method used for testing.
             */
            void TESTINGInitializer() {
                //Input how the data may be stored.
                inputData.Add("DEFENDANT_NAME", "John Doe");
                inputData.Add("DOB", "11/24/1995");
                inputData.Add("COUNT_NUMBER", "1");
                inputData.Add("OFFENSE_DATE", "3/8/2024");
                inputData.Add("CONDUCT", "Bad");
                inputData.Add("VICTIM", "Ronald Roe");

                //Print out the data stored in the dictionary.
                foreach (KeyValuePair<string, string> entry in inputData) {
                    Console.WriteLine($"{entry.Key}: {entry.Value}");
                }
            }

            //TESTINGInitializer();



            /* MESSING WITH CONTENT CONTROLS AND ACCESSING THEM
            Console.WriteLine(docRange.ContentControls.Count);
            Word.Range ctrl = docRange.ContentControls[1].Range;
            ctrl.Text = "testimber";
            */






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





            /*
            
            void findReplaceText() {

                //This foreach loops allows us to go through every 
                // key value pair, with the key being the word to replace
                // and the value being the replacement word.
                foreach (KeyValuePair<string, string> entry in inputData) {

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





            /*
            PRINTS OUT EVERY COLOR AVAILABLE

            for (int colRange = 1; colRange < 19; colRange++)
            {
                selection.Font.ColorIndex = (WdColorIndex)colRange;

                selection.TypeText($"{selection.Font.Color}\n");
            }

            



             CODE USED TO FIND A STRING AND REPLACE IT IN BOLD "you've been found"
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



            
                //Writes given text into the document.
            selection.TypeText("Hello to all!");

                //Saves the document at a specified directory.
            docx.SaveAs(directory);
            


            */
        }
    }
    
}