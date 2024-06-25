public class MultipleChoice
{
    // Properties
    public string Question { get; set; }
    public string[] Choices { get; set; }
    public string Answer { get; set; }

    // Constructor
    public MultipleChoice(string question, string[] choices, string answer)
    {
        Question = question;
        Choices = choices;
        Answer = answer;
    }

    public override string ToString()
    {
        string result = "";
        result += $"Q: {Question}\n";
        foreach (string choice in Choices)
        {
            result += $"{choice}\n";
        }
        result += $"A: {Answer}\n";

        return result;
    }
}