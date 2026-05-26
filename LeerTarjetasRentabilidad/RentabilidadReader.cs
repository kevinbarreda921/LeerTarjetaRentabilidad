using System;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using ExcelDataReader;

namespace LeerTarjetasRentabilidad;

public class RentabilidadReader
{
    public static void GenerarConfiguracionHojas()
    {
        // 1. Registrar proveedor de codificación para ExcelDataReader
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        string rutaProyecto = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
        string rutaArchivoExcel = Path.Combine(rutaProyecto, "ArchivosExcel", "Rentabilidad", "renabilidad vacio.xlsx");
        string rutaSalidaJson = Path.Combine(rutaProyecto, "ConfiguracionRentabilidad.json");

        Console.WriteLine($"\n[PROCESO] Iniciando lectura de hojas de rentabilidad...");
        Console.WriteLine($"Archivo de origen: {rutaArchivoExcel}");

        if (!File.Exists(rutaArchivoExcel))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n[ERROR] No se encontró el archivo Excel en la ruta especificada.");
            Console.WriteLine($"Ruta buscada: {rutaArchivoExcel}");
            Console.ResetColor();
            return;
        }

        try
        {
            var configHojas = new Dictionary<string, RangoHoja>();

            using (var stream = new FileStream(rutaArchivoExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                do
                {
                    string sheetName = reader.Name;
                    if (!string.IsNullOrWhiteSpace(sheetName))
                    {
                        // Ejemplo pre-configurado para demostración y prueba de duplicados
                        bool isCatacaos = sheetName.Equals("Catacaos", StringComparison.OrdinalIgnoreCase);
                        configHojas[sheetName] = new RangoHoja
                        {
                            FilaInicio = isCatacaos ? 135 : 0,
                            FilaFin = isCatacaos ? 145 : 0
                        };
                    }
                } while (reader.NextResult());
            }

            var options = new JsonSerializerOptions { WriteIndented = true };
            string jsonResultado = JsonSerializer.Serialize(configHojas, options);

            File.WriteAllText(rutaSalidaJson, jsonResultado);
            
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\n[ÉXITO] Se ha generado con éxito el archivo de configuración JSON en:");
            Console.WriteLine(rutaSalidaJson);
            Console.ResetColor();
            
            Console.WriteLine($"Total de hojas encontradas y configuradas: {configHojas.Count}");
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n[ERROR] Ocurrió un error inesperado al procesar el Excel: {ex.Message}");
            Console.ResetColor();
        }
    }
}

public class RangoHoja
{
    public int FilaInicio { get; set; }
    public int FilaFin { get; set; }
}
