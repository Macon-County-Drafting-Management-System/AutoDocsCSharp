
// See https://aka.ms/new-console-template for more information
Console.WriteLine("Hello, World!");


Dictionary<string, List<string>> inputData =
    new Dictionary<string, List<string>>();

void InputDataInitializer()
{
    inputData.Add("Defendent", ["Red", "John Doe"]);
}

InputDataInitializer();
foreach (KeyValuePair<string,List<string>> entry in inputData) {
    Console.WriteLine(entry);
}
