using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelDataReader;

namespace LeerTarjetasRentabilidad;

public class RentabilidadDiasReader
{
    public static void ProcesarDias()
    {
        // 1. Registrar proveedor de codificación para ExcelDataReader
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        string rutaProyecto = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\"));
        string rutaConfigJson = Path.Combine(rutaProyecto, "ConfiguracionRentabilidad.json");
        string rutaArchivoExcel = Path.Combine(rutaProyecto, "ArchivosExcel", "Rentabilidad", "renabilidad vacio.xlsx");
        string rutaSalidaJson = Path.Combine(rutaProyecto, "ResultadoDiasRentabilidad.json");

        Console.WriteLine($"\n[PROCESO] Iniciando extracción de días leídos...");
        
        if (!File.Exists(rutaConfigJson))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n[ERROR] No se encontró el archivo de configuración en: {rutaConfigJson}");
            Console.WriteLine("Por favor, ejecute primero la generación de la configuración.");
            Console.ResetColor();
            return;
        }

        if (!File.Exists(rutaArchivoExcel))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n[ERROR] No se encontró el archivo Excel en: {rutaArchivoExcel}");
            Console.ResetColor();
            return;
        }

        try
        {
            // 2. Cargar configuración de hojas y rangos
            string configText = File.ReadAllText(rutaConfigJson);
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            var configHojas = JsonSerializer.Deserialize<Dictionary<string, RangoHoja>>(configText, options);

            if (configHojas == null || configHojas.Count == 0)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("\n[ADVERTENCIA] La configuración de hojas está vacía o no es válida.");
                Console.ResetColor();
                return;
            }

            var resultadoFinal = new Dictionary<string, List<RegistroDia>>();

            using (var stream = new FileStream(rutaArchivoExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                // Iterar sobre cada hoja del Excel
                do
                {
                    string sheetName = reader.Name;

                    // Buscar si la hoja actual está en nuestra configuración
                    var configKey = configHojas.Keys.FirstOrDefault(k => k.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
                    if (configKey == null) continue;

                    var config = configHojas[configKey];

                    // Si FilaInicio es 0 o no está configurada, omitimos el procesamiento de filas de esta hoja
                    if (config.FilaInicio <= 0 || config.FilaFin <= 0 || config.FilaInicio > config.FilaFin)
                    {
                        continue;
                    }

                    Console.WriteLine($"[PROCESO] Leyendo hoja '{sheetName}' (Filas {config.FilaInicio} a {config.FilaFin})...");

                    var registrosHoja = new List<RegistroDia>();
                    int filaActual = 1;

                    while (reader.Read())
                    {
                        if (filaActual >= config.FilaInicio && filaActual <= config.FilaFin)
                        {
                            int colCount = reader.FieldCount;

                            // Leer la fecha de la columna index 1 (columna B)
                            object? rawFecha = colCount > 1 ? reader.GetValue(1) : null;
                            string? fechaFormatted = FormatearFecha(rawFecha);

                            if (!string.IsNullOrWhiteSpace(fechaFormatted))
                            {
                                registrosHoja.Add(new RegistroDia
                                {
                                    FechaHoja = fechaFormatted,
                                    Fila = filaActual
                                });
                            }
                        }

                        if (filaActual > config.FilaFin)
                        {
                            break;
                        }

                        filaActual++;
                    }

                    // Regla: "solo deja el primer registro fecha repetido"
                    // Filtramos por fecha única quedándonos con el primer registro encontrado
                    var registrosFiltrados = registrosHoja
                        .GroupBy(r => r.FechaHoja)
                        .Select(g => g.First())
                        .OrderBy(r => r.Fila) // Mantener orden ascendente por fila
                        .ToList();

                    if (registrosFiltrados.Count > 0)
                    {
                        resultadoFinal[sheetName] = registrosFiltrados;
                    }

                } while (reader.NextResult());
            }

            // 3. Serializar y guardar el nuevo archivo JSON
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string jsonResultado = JsonSerializer.Serialize(resultadoFinal, jsonOptions);

            File.WriteAllText(rutaSalidaJson, jsonResultado);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\n[ÉXITO] Se ha generado con éxito el archivo JSON de días leídos en:");
            Console.WriteLine(rutaSalidaJson);
            Console.ResetColor();

            Console.WriteLine($"Total de hojas procesadas con datos válidos: {resultadoFinal.Count}");
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n[ERROR] Ocurrió un error al extraer los días: {ex.Message}");
            Console.ResetColor();
        }
    }

    #region Helpers de Conversión y Formato

    // Formatea la fecha de Excel de forma limpia en formato "d/MM/yyyy"
    private static string? FormatearFecha(object? rawFecha)
    {
        if (rawFecha == null) return null;
        if (rawFecha is DateTime dt) return dt.ToString("d/MM/yyyy");
        
        if (DateTime.TryParse(rawFecha.ToString(), out DateTime parsedDate)) 
            return parsedDate.ToString("d/MM/yyyy");

        return null;
    }

    #endregion
}

public class RegistroDia
{
    [JsonPropertyName("fecha_hoja")]
    public string? FechaHoja { get; set; }

    [JsonPropertyName("Fila")]
    public int Fila { get; set; }
}
