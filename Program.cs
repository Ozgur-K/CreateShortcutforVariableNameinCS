// See https://aka.ms/new-console-template for more information
using System.Diagnostics;

//Console.WriteLine("Hello, World!");

new Shortcut(){ };

class Shortcut
{

    public String folderName;
    public String[] files;


    public Shortcut()
    {

        folderName = "C:\\Users\\ozgur\\Desktop\\";
        button1_Click();

    }

    private String FindExcelFileandGetpath()
    {
        files = Directory.GetFiles(folderName);
        String theFile = "null";

        foreach (String file in files)
        {
            if (File.Exists(file))
            {
                Debug.WriteLine(file);
                if (file.Contains("picture"))
                {
                    theFile = file;
                }
            }
        }
        return theFile;
    }


    private void button1_Click()
    {
        // OpenWithDialog();
        FindExcelFileandGetpath();
        Debug.WriteLine(FindExcelFileandGetpath());
        OpenWithDefaultProgram(FindExcelFileandGetpath());
    }

    public void OpenWithDefaultProgram(string path)
    {
        using Process fileopener = new Process();

        fileopener.StartInfo.FileName = "explorer";
        fileopener.StartInfo.Arguments = "\"" + path + "\"";
        fileopener.Start();
    }

    private void OpenWithDialog()
    {
        var args = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "shell32.dll");
        args += ",OpenAs_RunDLL " + FindExcelFileandGetpath();
        Process.Start("rundll32.exe", args);
    }
}
    
