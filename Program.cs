using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using OfficeOpenXml;
using Microsoft.InformationProtection;
using Microsoft.InformationProtection.File;
using Microsoft.InformationProtection.Protection;

namespace epplus
{
    internal class Program
    {
        // Reemplaza estos valores con los de tu aplicación
        private const string ClientId = "<CLIENT_ID>";
        private const string AppName = "DemoEPPlusMIP";
        private const string UserEmail = "<USUARIO>";      // p.ej. usuario@dominio.com
        private const string LabelId = "<LABEL_ID>";      // ID de la etiqueta de sensibilidad

        static void Main(string[] args)
        {
            // Aceptar la licencia de EPPlus
            // EPPlus LicenseContext: https://epplussoftware.com/docs/5.0/articles/licensing.html
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Inicializar SDK MIP para operaciones de archivo
            // Doc MIP.Initialize: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#initialize-the-sdk
            MIP.Initialize(MipComponent.File);

            // Configurar ApplicationInfo
            // Doc ApplicationInfo: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#define-applicationinfo
            var appInfo = new ApplicationInfo
            {
                ApplicationId = ClientId,
                ApplicationName = AppName,
                ApplicationVersion = "1.0.0"
            };

            // Instanciar los delegados de autenticación y consentimiento
            // Doc AuthDelegate: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#authentication
            var authDelegate = new AuthDelegateImplementation(appInfo);
            // Doc ConsentDelegate: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#consent
            var consentDelegate = new ConsentDelegateImplementation();

            // Crear configuración y contexto de MIP
            // Doc MipConfiguration y CreateMipContext: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#create-mipcontext
            var mipConfiguration = new MipConfiguration(appInfo, "mip_data", LogLevel.Trace, false);
            var mipContext = MIP.CreateMipContext(mipConfiguration);

            // Configurar y cargar el perfil de archivo
            // Doc FileProfileSettings y LoadFileProfileAsync: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#load-file-profile
            var profileSettings = new FileProfileSettings(mipContext, CacheStorageType.OnDiskEncrypted, consentDelegate);
            var fileProfile = Task.Run(async () => await MIP.LoadFileProfileAsync(profileSettings)).Result;

            // Configurar y crear el motor de archivo
            // Doc FileEngineSettings y AddEngineAsync: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#create-engine
            var engineSettings = new FileEngineSettings(UserEmail, authDelegate, "", "es-ES");
            engineSettings.Identity = new Identity(UserEmail);
            var fileEngine = Task.Run(async () => await fileProfile.AddEngineAsync(engineSettings)).Result;

            // Crear datos de ejemplo y generar archivo Excel
            var dt = new DataTable("Empleados");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Nombre", typeof(string));
            dt.Columns.Add("Departamento", typeof(string));
            dt.Rows.Add(1, "Juan Pérez", "IT");
            dt.Rows.Add(2, "Ana López", "Finanzas");
            dt.Rows.Add(3, "Carlos Ruiz", "Recursos Humanos");

            using (var package = new ExcelPackage())
            {
                var ws = package.Workbook.Worksheets.Add("Datos");
                ws.Cells["A1"].LoadFromDataTable(dt, true);

                var inputFilePath = Path.Combine(Environment.CurrentDirectory, "Empleados.xlsx");
                package.SaveAs(new FileInfo(inputFilePath));
                Console.WriteLine($"Archivo Excel generado: {inputFilePath}");

                // Aplicar etiqueta de sensibilidad
                ApplySensitivityLabel(fileEngine, inputFilePath);
            }

            // Liberar recursos y cerrar contexto MIP
            // Doc ShutDown: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#shutdown
            fileProfile = null;
            fileEngine = null;
            mipContext.ShutDown();
            mipContext = null;
        }

        static void ApplySensitivityLabel(IFileEngine fileEngine, string inputFilePath)
        {
            // Doc CreateFileHandlerAsync: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#create-filehandler
            var handler = Task.Run(async () =>
                await fileEngine.CreateFileHandlerAsync(inputFilePath, inputFilePath, true)).Result;

            // Obtener la etiqueta por ID
            // Doc GetLabelById: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#get-labels
            var label = fileEngine.GetLabelById(LabelId);
            var labelingOptions = new LabelingOptions
            {
                AssignmentMethod = AssignmentMethod.Standard
            };

            // Aplicar la etiqueta con configuración de protección predeterminada
            // Doc SetLabel: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#apply-label
            handler.SetLabel(label, labelingOptions, new ProtectionSettings());

            // Confirmar los cambios y sobrescribir el archivo
            // Doc CommitAsync: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#commit
            var result = Task.Run(async () =>
                await handler.CommitAsync(inputFilePath)).Result;

            Console.WriteLine($"Etiqueta '{label.Name}' aplicada al archivo.");
        }
    }
}
