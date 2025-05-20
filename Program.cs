// Target .NET Framework 4.8
using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using OfficeOpenXml;
using Microsoft.Identity.Client;
using Microsoft.InformationProtection;
using Microsoft.InformationProtection.File;
using Microsoft.InformationProtection.Protection;

namespace epplus
{
    internal class Program
    {
        // Reemplaza estos valores con los de tu aplicación
        private const string ClientId = "<CLIENT_ID>";
        private const string ClientSecret = "<CLIENT_SECRET>";
        private const string TenantId = "<TENANT_ID>";
        private const string AppName = "DemoEPPlusMIP";
        private const string UserEmail = "<USUARIO>";      // p.ej. usuario@dominio.com
        private const string LabelId = "<LABEL_ID>";      // ID de la etiqueta de sensibilidad

        static void Main(string[] args)
        {
            // 1. Aceptar la licencia de EPPlus
            // https://epplussoftware.com/docs/5.0/articles/licensing.html
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 2. Inicializar SDK MIP para operaciones de archivo
            // https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#initialize-the-sdk
            MIP.Initialize(MipComponent.File);

            // 3. Configurar ApplicationInfo
            // https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#define-applicationinfo
            var appInfo = new ApplicationInfo
            {
                ApplicationId = ClientId,
                ApplicationName = AppName,
                ApplicationVersion = "1.0.0"
            };

            // 4. Instanciar los delegados de autenticación y consentimiento
            // AuthDelegateImplementation (Client Credentials)
            // https://learn.microsoft.com/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
            var authDelegate = new AuthDelegateImplementation(ClientId, ClientSecret, TenantId);
            // ConsentDelegateImplementation
            // https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#consent-delegate
            var consentDelegate = new ConsentDelegateImplementation();

            // 5. Crear configuración y contexto de MIP
            // https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#create-mipcontext
            var mipConfiguration = new MipConfiguration(appInfo, "mip_data", LogLevel.Trace, false);
            var mipContext = MIP.CreateMipContext(mipConfiguration);

            // 6. Configurar y cargar el perfil de archivo
            // https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#load-file-profile
            var profileSettings = new FileProfileSettings(mipContext, CacheStorageType.OnDiskEncrypted, consentDelegate);
            var fileProfile = Task.Run(async () => await MIP.LoadFileProfileAsync(profileSettings)).Result;

            // 7. Configurar y crear el motor de archivo
            // https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#create-engine
            var engineSettings = new FileEngineSettings(UserEmail, authDelegate, string.Empty, "es-ES")
            {
                Identity = new Identity(UserEmail)
            };
            var fileEngine = Task.Run(async () => await fileProfile.AddEngineAsync(engineSettings)).Result;

            // 8. Crear datos de ejemplo y generar archivo Excel
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

                // 9. Aplicar etiqueta de sensibilidad
                ApplySensitivityLabel(fileEngine, inputFilePath);
            }

            // 10. Liberar recursos y cerrar contexto MIP
            // https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#shutdown
            fileProfile = null;
            fileEngine = null;
            mipContext.ShutDown();
            mipContext = null;
        }

        static void ApplySensitivityLabel(IFileEngine fileEngine, string inputFilePath)
        {
            // CreateFileHandlerAsync: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#create-filehandler
            var handler = Task.Run(async () => await fileEngine.CreateFileHandlerAsync(inputFilePath, inputFilePath, true)).Result;

            // GetLabelById: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#get-labels
            var label = fileEngine.GetLabelById(LabelId);
            var labelingOptions = new LabelingOptions { AssignmentMethod = AssignmentMethod.Standard };

            // SetLabel con ProtectionSettings: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#apply-label
            handler.SetLabel(label, labelingOptions, new ProtectionSettings());

            // CommitAsync: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#commit
            Task.Run(async () => await handler.CommitAsync(inputFilePath)).Wait();
            Console.WriteLine($"Etiqueta '{label.Name}' aplicada al archivo.");
        }
    }

    /// <summary>
    /// Implementación de IAuthDelegate usando MSAL y Client Credentials
    /// </summary>
    public class AuthDelegateImplementation : IAuthDelegate
    {
        private readonly IConfidentialClientApplication _clientApp;

        /// <summary>
        /// Constructor para flujo de credenciales de cliente
        /// </summary>
        public AuthDelegateImplementation(string clientId, string clientSecret, string tenantId)
        {
            // Confidential Client App MSAL: https://learn.microsoft.com/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow
            _clientApp = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithTenantId(tenantId)
                .Build();
        }

        /// <summary>
        /// AcquireTokenForClient: https://learn.microsoft.com/dotnet/api/microsoft.identity.client.confidentialclientapplicationbuilder
        /// </summary>
        public async Task<string> AcquireTokenAsync(IEnumerable<string> scopes)
        {
            var result = await _clientApp.AcquireTokenForClient(scopes).ExecuteAsync();
            return result.AccessToken;
        }
    }

    /// <summary>
    /// Implementación básica de IConsentDelegate
    /// </summary>
    public class ConsentDelegateImplementation : IConsentDelegate
    {
        // ConsentDelegate: https://learn.microsoft.com/information-protection/develop/quick-file-set-get-label-csharp#consent-delegate
        public Task<ConsentResult> ShouldShowConsentAsync(ConsentCallback callback, ConsentType consentType, string resourceId, string resourceName)
        {
            return Task.FromResult(ConsentResult.Create(ConsentDecision.Yes, DateTimeOffset.UtcNow));
        }
    }
}
