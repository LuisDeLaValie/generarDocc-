using System;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace EditarDocumentoWord
{
    class Program
    {
        static void Main(string[] args)
        {
            // Ruta del documento de Word existente
            string filePath = @"C:\Users\TDTx\Documents\generarDoc\documento.docx";

            // Verificar si el archivo existe
            if (!File.Exists(filePath))
            {
                Console.WriteLine("El archivo no existe.");
                return;
            }

            // Crear una instancia de la aplicación Word
            Application wordApp = new Application();

            // Abrir el documento existente
            Document doc = wordApp.Documents.Open(filePath);

            // Obtener el contenido del documento
            string docContent = doc.Content.Text;

            // Generar HTML a partir del contenido del documento
            string htmlContent = "<html><body>" + docContent + "</body></html>";

            // Realizar cambios en el contenido HTML (aquí puedes editar el HTML como desees)

            // Por ejemplo, agregar un nuevo párrafo
            htmlContent += "<p>Este es un nuevo párrafo agregado mediante C#.</p>";

            // Limpiar el contenido anterior del documento
            doc.Content.Delete();

            // Insertar el contenido HTML modificado en el documento
            doc.Content.InsertXML(htmlContent);

            // Guardar los cambios
            doc.Save();

            // Cerrar el documento
            doc.Close();

            // Cerrar la aplicación Word
            wordApp.Quit();

            Console.WriteLine("Documento de Word editado exitosamente con contenido HTML.");
            Console.ReadLine();
        }
    }
}
