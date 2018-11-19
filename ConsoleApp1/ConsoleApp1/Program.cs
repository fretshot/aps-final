using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.IO;


namespace ConsoleApp1{
    class Program {
        static void Main(string[] args) {

            Document doc = new Document(PageSize.LETTER);
            PdfWriter writer = PdfWriter.GetInstance(doc,new FileStream(@"C:\Users\oscar\Desktop\aps\software.pdf", FileMode.Create, FileAccess.ReadWrite));

            doc.AddTitle("Software Instalado");
            doc.AddCreator("Administrador");

            doc.Open();
            iTextSharp.text.Font font = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);

            doc.Add(new Paragraph("*** DEPARTAMENTO DE SISTEMAS ***"));
            doc.Add(Chunk.NEWLINE);
            doc.Add(new Paragraph("Gestión de Software"));
            doc.Add(Chunk.NEWLINE);
            doc.Add(new Paragraph("Fecha: " + DateTime.Now.ToString("d/M/yyyy")));
            doc.Add(new Paragraph("Maquina: "+ Environment.MachineName));
            doc.Add(new Paragraph("Usuario: " + Environment.UserName));
            doc.Add(new Paragraph("Sistema: " + Environment.OSVersion.VersionString));
            doc.Add(Chunk.NEWLINE);
            
            string registry_key = @"Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\";
            using (Microsoft.Win32.RegistryKey key = Registry.LocalMachine.OpenSubKey(registry_key)) {
                foreach (string subkey_name in key.GetSubKeyNames()) {
                    using (RegistryKey subkey = key.OpenSubKey(subkey_name)) {
                        string s = JsonConvert.SerializeObject(subkey.GetValue("DisplayName"));
                        if (s != "null") {
                            doc.Add(new Paragraph(s,font));
                        }
                    }
                }
            }
            doc.Add(Chunk.NEWLINE);
            doc.Add(Chunk.NEWLINE);
            doc.Add(new Paragraph("Firma de conformidad: ______________________________________________"));
            doc.Add(Chunk.NEWLINE);
            doc.Close();
            writer.Close();
            Console.WriteLine("Documento creado!\n");
            Console.WriteLine("Presione cualquier tecla para continuar...");
            Console.ReadKey();
        }
    }
}
