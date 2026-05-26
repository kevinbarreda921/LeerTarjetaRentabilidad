using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using ExcelDataReader;

namespace LeerTarjetasRentabilidad;

public class TarjetasProcesador
{
    public static void Procesar()
    {
        // 1. Registrar proveedor de codificación para ExcelDataReader
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        string rutaProyecto = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
        string rutaArchivoExcel = Path.Combine(rutaProyecto, "ArchivosExcel", "ResumenTarjeta", "resumen tarjetas.xlsx");
        string rutaArchivoJson = Path.Combine(rutaProyecto, "GrifosLeer.json");
        string rutaSalidaJson = Path.Combine(rutaProyecto, "ResultadoGrifos.json");

        try
        {
            // 1. Definir la lista de ejecuciones (hojas y rangos de filas a procesar)
            var ejecuciones = new List<(string SheetName, int StartRow, int EndRow)>
            {
                ("Lou", 80, 110),
                ("Lou", 119, 148),
                ("Coco", 80, 110),
                ("Coco", 118, 147),
                ("Jeaneth", 234, 264),
                ("Jeaneth", 272, 301),
                ("Mavel completo", 298, 328),
                ("Mavel completo", 335, 364),
                ("Juan Carlos", 160, 190),
                ("Juan Carlos", 198, 227)
            };

            var registros = new List<RegistroTarjeta>();

            Console.WriteLine("Iniciando procesamiento de hojas y rangos...");

            // 2. Ejecutar cada una en un bucle y acumular los registros
            foreach (var (sheetName, startRow, endRow) in ejecuciones)
            {
                Console.WriteLine($"\n[PROCESO] Cargando configuración para la hoja '{sheetName}'...");
                var configGrifos = CargarConfiguracionGrifos(rutaArchivoJson, sheetName);

                if (configGrifos == null || configGrifos.Count == 0)
                {
                    MostrarError($"No se encontró configuración para la hoja '{sheetName}' en el archivo JSON.");
                    continue; // Continuar con las demás ejecuciones si alguna falla
                }

                Console.WriteLine($"[PROCESO] Leyendo Excel - Hoja: '{sheetName}' | Filas: {startRow} a {endRow}...");
                var registrosHoja = LeerExcelDinamico(rutaArchivoExcel, sheetName, configGrifos, startRow, endRow);
                registros.AddRange(registrosHoja);
            }

            Console.WriteLine($"\n[PROCESO] Lectura completada. Se procesaron {registros.Count} registros en total.");

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

            Console.WriteLine("\n--- RESUMEN DE PROCESAMIENTO ---");
            Console.WriteLine($"Total de registros procesados: {registros.Count}");
            Console.WriteLine($"Total de grifos agrupados: {agrupadoPorGrifo.Count}");
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
    }

    // Helper para mostrar errores en consola en color rojo
    private static void MostrarError(string mensaje)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"\n[ERROR] {mensaje}");
        Console.ResetColor();
    }

    public static Dictionary<string, GrifoConfig>? CargarConfiguracionGrifos(string rutaJson, string sheetName) 
        => ExcelService.CargarConfiguracionGrifos(rutaJson, sheetName);

    public static List<RegistroTarjeta> LeerExcelDinamico(string filePath, string sheetName, Dictionary<string, GrifoConfig> grifosConfig, int startRow, int endRow)
        => ExcelService.LeerExcelDinamico(filePath, sheetName, grifosConfig, startRow, endRow);
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
