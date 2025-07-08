namespace WordToPDF;

class Program
{
    static void Main(string[] args)
    {
        //Here you must enter the file path, that means it should have something like C:\Users\{User}\Desktop\file.docx
        //Aquí se debe ingresar la ruta del archivo, eso significa que debería tener algo tipo C:\Users\{User}\Desktop\file.docx
        string DocxPath = @"C:\{your_path_here}";
        //Here you have to enter the output path of the file, that means it should have something like C:\{User}\{Desktop}.
        //Aquí se debe ingresar la ruta de salida del archivo, eso significa que debería tener algo tipo C:\Users\{User}\Desktop
        string OutPutFolder = @"C:\{your_path_here}";
        ConvertDocxToPdf(DocxPath, OutPutFolder);


    }

    public static void ConvertDocxToPdf(string docxPath, string outputFolder)
    {
        //Here you must enter the path where you have the portable LibreOffice folder
        //Aqui debe ingresar la ruta donde tiene la carpeta portable de LibreOffice
        string libreOfficePath = @"C:\Users\{User}\Desktop\LibreOfficePortable\App\libreoffice\program\soffice.exe";

        ProcessStartInfo startInfo = new()
        {
            FileName = libreOfficePath,
            Arguments = $"--headless --convert-to pdf \"{docxPath}\" --outdir \"{outputFolder}\"", // Specifies the command-line arguments to pass to the program.
            RedirectStandardOutput = true, //Allows you to redirect the process’s output stream.
            RedirectStandardError = true, //Allows you to redirect the process’s error stream.
            UseShellExecute = false, //Indicates whether to use the operating system shell to start the process.
            CreateNoWindow = true // Specifies whether to create a new window for the process.
        };

        using Process process = new() { StartInfo = startInfo };
        process.Start();

        string output = process.StandardOutput.ReadToEnd();
        string error = process.StandardError.ReadToEnd();

        process.WaitForExit();

        if (process.ExitCode == 0)
        {
            Console.WriteLine("Conversión completada.");
        }
        else
        {
            Console.WriteLine("Error al convertir:");
            Console.WriteLine(error);
        }
        //rename pdf file
        string archivoOriginal = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docxPath) + ".pdf");
        string archivoRenombrado = Path.Combine(outputFolder, "nuevoNombre.pdf");
        
        if (File.Exists(archivoOriginal))
        {
            File.Move(archivoOriginal, archivoRenombrado);
            Console.WriteLine($"Archivo renombrado a {archivoRenombrado}");
        }
    }
}
