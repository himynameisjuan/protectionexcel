using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using OfficeOpenXml;
using Microsoft.InformationProtection;
using Microsoft.InformationProtection.File;
using System.ComponentModel;

namespace epplus
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Aceptar la licencia de EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Crear una tabla ficticia
            var dt = new DataTable("Empleados");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Nombre", typeof(string));
            dt.Columns.Add("Departamento", typeof(string));
            dt.Rows.Add(1, "Juan Pérez", "IT");
            dt.Rows.Add(2, "Ana López", "Finanzas");
            dt.Rows.Add(3, "Carlos Ruiz", "Recursos Humanos");

            // Crear el archivo Excel
            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Datos");
                ws.Cells["A1"].LoadFromDataTable(dt, true);

                // Guardar el archivo temporalmente
                var filePath = Path.Combine(Environment.CurrentDirectory, "Empleados.xlsx");
                package.SaveAs(new FileInfo(filePath));

                Console.WriteLine($"Archivo Excel generado: {filePath}");

                // Aplicar la Sensitivity Label usando MIP SDK
                ApplySensitivityLabel(filePath);
            }
        }

        static void ApplySensitivityLabel(string filePath)
        {
            // Configuración de la aplicación
            var appInfo = new ApplicationInfo()
            {
                ApplicationId = "<CLIENT_ID>",
                ApplicationName = "DemoEPPlusMIP",
                ApplicationVersion = "1.0.0"
            };

            // Inicializar MIP SDK
            MIP.Initialize(MipComponent.File);

            // Crear el profile de MIP usando FileProfileSettings
            var profileSettings = new FileProfileSettings(@"C:\Logs", LogLevel.Trace);
            var profile = MIP.LoadFileProfileAsync(appInfo, profileSettings).Result;

            // Autenticación (implementa tu propio método de autenticación)
            var authDelegate = new AuthDelegateImplementation("<CLIENT_ID>", "<CLIENT_SECRET>", "<TENANT_ID>");

            // Crear el engine
            var engineSettings = new FileEngineSettings("<USUARIO>", "es-ES", authDelegate, "", "");
            var engine = profile.AddEngineAsync(engineSettings).Result;

            // Obtener la etiqueta de sensibilidad (por ejemplo, la primera disponible)
            var labels = engine.SensitivityLabels;
            var label = labels[0]; // O busca por nombre/ID

            // Cargar el archivo y aplicar la etiqueta
            var fileHandler = engine.CreateFileHandlerAsync(filePath, FileAccessMode.ReadWrite, true).Result;
            fileHandler.SetLabel(label, new LabelingOptions() { AssignmentMethod = AssignmentMethod.Standard });
            fileHandler.CommitAsync(filePath).Wait();

            Console.WriteLine($"Etiqueta de sensibilidad '{label.Name}' aplicada al archivo.");
        }
    }
}
