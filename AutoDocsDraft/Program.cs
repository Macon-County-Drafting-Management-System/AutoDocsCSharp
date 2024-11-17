using Word = Microsoft.Office.Interop.Word;

string directory = "C:\\users\\shuff\\documents\\csharpdocs\\CSharpDocTest.docx";

var wordApp = new Word.Application();
wordApp.Visible = true;

var docx = wordApp.Documents.Add();
var selection = wordApp.Selection;
selection.TypeText("Hello to all!");
docx.SaveAs2(directory);


Dictionary<string, List<string>> inputData = new();

void TESTINGInitializer(){
    inputData.Add("Red", ["Defendent", "John Doe"]);
}

TESTINGInitializer();
foreach (KeyValuePair<string,List<string>> entry in inputData) {
    Console.WriteLine($"{entry.Key}: {entry.Value[0]}, {entry.Value[1]}");
}
