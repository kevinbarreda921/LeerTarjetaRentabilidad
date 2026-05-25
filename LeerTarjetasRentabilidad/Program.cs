using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using ExcelDataReader;

// 1. Registrar proveedor de codificación para ExcelDataReader
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

string rutaProyecto = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
string rutaArchivoExcel = Path.Combine(rutaProyecto, "ArchivosExcel", "ResumenTarjeta", "resumen tarjetas.xlsx");
string rutaArchivoJson = Path.Combine(rutaProyecto, "GrifosLeer.json");
string rutaSalidaJson = Path.Combine(rutaProyecto, "ResultadoGrifos.json");

try
{
    string sheetName = "Juan Carlos"; // Puede ser "Lou", "Coco", etc.

    Console.WriteLine($"Cargando la configuración de grifos para la hoja '{sheetName}' desde JSON...");
    var configGrifos = CargarConfiguracionGrifos(rutaArchivoJson, sheetName);

    if (configGrifos == null || configGrifos.Count == 0)
    {
        MostrarError($"No se encontró configuración para la hoja '{sheetName}' en el archivo JSON.");
        return;
    }

    Console.WriteLine($"Configuración cargada con éxito. Se detectaron {configGrifos.Count} grifos a procesar.");
    Console.WriteLine("Iniciando la lectura del archivo Excel...");

    // 2. Leer los registros dinámicamente según la configuración
    //List<RegistroTarjeta> registros = LeerExcelDinamico(rutaArchivoExcel, "Lou", configGrifos, 80, 110);
    //List<RegistroTarjeta> registros = LeerExcelDinamico(rutaArchivoExcel, "Lou", configGrifos, 119, 148);
    //List<RegistroTarjeta> registros = LeerExcelDinamico(rutaArchivoExcel, "Coco", configGrifos, 80, 110);
    //List<RegistroTarjeta> registros = LeerExcelDinamico(rutaArchivoExcel, "Coco",configGrifos, 118, 147);
    //List<RegistroTarjeta> registros = LeerExcelDinamico(rutaArchivoExcel, "Jeaneth", configGrifos, 234, 264);
    //List<RegistroTarjeta> registros = LeerExcelDinamico(rutaArchivoExcel, "Jeaneth", configGrifos, 272, 301);
    //List<RegistroTarjeta> registros = LeerExcelDinamico(rutaArchivoExcel, "Mavel completo", configGrifos, 298, 328);
    //List<RegistroTarjeta> registros = LeerExcelDinamico(rutaArchivoExcel, "Mavel completo", configGrifos, 335, 364);
    List<RegistroTarjeta> registros = LeerExcelDinamico(rutaArchivoExcel, "Juan Carlos", configGrifos, 160, 190);
    //List<RegistroTarjeta> registros = LeerExcelDinamico(rutaArchivoExcel, "Juan Carlos", configGrifos, 198, 227);

    // 3. Agrupar los resultados por grifo en el formato solicitado
    var agrupadoPorGrifo = registros
        .GroupBy(r => r.Grifo)
        .ToDictionary(
            grupo => grupo.Key,
            grupo => grupo.Select(r => new
            {
                fecha_hoja = r.FechaHoja,
                DataLiquidos = r.Liquidos,
                DataDescuento = r.Descuento
            }).ToList()
        );

    // 4. Serializar a un nuevo archivo JSON formateado
    var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
    string jsonResultado = JsonSerializer.Serialize(agrupadoPorGrifo, jsonOptions);

    File.WriteAllText(rutaSalidaJson, jsonResultado);
    Console.WriteLine($"\n[ÉXITO] Se ha generado con éxito el archivo JSON en:");
    Console.WriteLine(rutaSalidaJson);

    Console.WriteLine("\n--- JSON GENERADO ---");
    Console.WriteLine(jsonResultado);


}
catch (FileNotFoundException ex)
{
    MostrarError($"Error: No se encontró un archivo requerido. Detalle: {ex.Message}");
}
catch (InvalidOperationException ex)
{
    MostrarError($"Error de Operación: {ex.Message}");
}
catch (Exception ex)
{
    MostrarError($"Ocurrió un error inesperado: {ex.Message}");
}

// Helper para mostrar errores en consola en color rojo
static void MostrarError(string mensaje)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine($"\n[ERROR] {mensaje}");
    Console.ResetColor();
}

#region Modelos y Clases de Configuración

/// <summary>
/// Representa la configuración de celdas para un grifo específico.
/// </summary>
public class GrifoConfig
{
    public int celda_liquidos { get; set; }
    public int celda_Descuento { get; set; }
}

/// <summary>
/// Representa el registro final leído y procesado de la tarjeta para un grifo y fecha.
/// </summary>
public record RegistroTarjeta(
    string Grifo,
    int Fila,
    string? FechaHoja,
    decimal Liquidos,
    decimal Descuento
);

#endregion

#region Servicios de Excel y JSON

public static class ExcelService
{
    // Carga y deserializa el JSON de configuración para una hoja específica
    public static Dictionary<string, GrifoConfig>? CargarConfiguracionGrifos(string rutaJson, string sheetName)
    {
        if (!File.Exists(rutaJson))
        {
            throw new FileNotFoundException("El archivo JSON de configuración no existe.", rutaJson);
        }

        string jsonText = File.ReadAllText(rutaJson);
        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        
        var fullConfig = JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, GrifoConfig>>>(jsonText, options);

        if (fullConfig != null && fullConfig.TryGetValue(sheetName, out var grifosConfig))
        {
            return grifosConfig;
        }

        return null;
    }

    // Lee el archivo Excel mapeando dinámicamente las columnas de los grifos especificados
    public static List<RegistroTarjeta> LeerExcelDinamico(
        string filePath,
        string sheetName,
        Dictionary<string, GrifoConfig> grifosConfig,
        int startRow,
        int endRow)
    {
        var campofecha=1;
        if (sheetName == "Juan Carlos")
        {
            campofecha = 0;
        }

        var resultList = new List<RegistroTarjeta>();

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException("El archivo Excel especificado no existe.", filePath);
        }

        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 4096, FileOptions.SequentialScan);
        using var reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

        bool sheetFound = false;

        do
        {
            if (reader.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
            {
                sheetFound = true;
                int filaActual = 1;

                while (reader.Read())
                {
                    if (filaActual >= startRow && filaActual <= endRow)
                    {
                        int colCount = reader.FieldCount;

                        // La fecha de la hoja siempre se encuentra en la columna index 1 (columna B)
             
                        object? rawFecha = colCount > 1 ? reader.GetValue(campofecha) : null;
                        string? fechaHoja = FormatearFecha(rawFecha);

                        // Iterar por cada grifo configurado en el JSON e insertar en la lista
                        foreach (var kvp in grifosConfig)
                        {
                            string grifoNombre = kvp.Key;
                            GrifoConfig config = kvp.Value;

                            object? rawLiquidos = colCount > config.celda_liquidos ? reader.GetValue(config.celda_liquidos) : null;
                            object? rawDescuento = colCount > config.celda_Descuento ? reader.GetValue(config.celda_Descuento) : null;

                            decimal liquidos = ConvertirDosDecimales(rawLiquidos);
                            decimal descuento = ConvertirDosDecimales(rawDescuento);

                            resultList.Add(new RegistroTarjeta(
                                Grifo: grifoNombre,
                                Fila: filaActual,
                                FechaHoja: fechaHoja,
                                Liquidos: liquidos,
                                Descuento: descuento
                            ));
                        }
                    }

                    if (filaActual > endRow)
                    {
                        break;
                    }

                    filaActual++;
                }

                return resultList;
            }
        } while (reader.NextResult());

        if (!sheetFound)
        {
            throw new InvalidOperationException($"La hoja '{sheetName}' no fue encontrada en el archivo Excel.");
        }

        return resultList;
    }

    #region Helpers de Conversión y Formato

    // Convierte un objeto a formato decimal redondeado a 2 posiciones
    private static decimal ConvertirDosDecimales(object? val)
    {
        if (val == null) return 0m;
        if (val is decimal dec) return Math.Round(dec, 2);
        if (val is double d) return Math.Round((decimal)d, 2);
        if (val is float f) return Math.Round((decimal)f, 2);
        if (val is int i) return (decimal)i;

        if (decimal.TryParse(val.ToString(), out decimal parsed))
        {
            return Math.Round(parsed, 2);
        }

        return 0m;
    }

    // Formatea la fecha de Excel de forma limpia en formato "d/MM/yyyy"
    private static string? FormatearFecha(object? rawFecha)
    {
        if (rawFecha == null) return null;
        if (rawFecha is DateTime dt) return dt.ToString("d/MM/yyyy");
        if (DateTime.TryParse(rawFecha.ToString(), out DateTime parsedDate)) return parsedDate.ToString("d/MM/yyyy");

        string rawStr = rawFecha.ToString() ?? "";
        int indexSpace = rawStr.IndexOf(" 00:00");
        return indexSpace != -1 ? rawStr.Substring(0, indexSpace) : rawStr.Trim();
    }

    #endregion
}

#endregion

// Exponer métodos de ExcelService de forma global para simplificar el flujo superior
public partial class Program
{
    public static Dictionary<string, GrifoConfig>? CargarConfiguracionGrifos(string rutaJson, string sheetName) 
        => ExcelService.CargarConfiguracionGrifos(rutaJson, sheetName);

    public static List<RegistroTarjeta> LeerExcelDinamico(string filePath, string sheetName, Dictionary<string, GrifoConfig> grifosConfig, int startRow, int endRow)
        => ExcelService.LeerExcelDinamico(filePath, sheetName, grifosConfig, startRow, endRow);
}