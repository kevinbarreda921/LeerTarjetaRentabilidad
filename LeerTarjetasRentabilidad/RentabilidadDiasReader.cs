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
        string rutaArchivoExcel = Path.Combine(rutaProyecto, "ArchivosExcel", "Rentabilidad", "renabilidad vacio.xlsx");
        string rutaSalidaJson = Path.Combine(rutaProyecto, "ResultadoDiasRentabilidad.json");

        Console.WriteLine($"\n[PROCESO] Iniciando escaneo automático de días leídos (Marzo y Abril 2026)...");
        Console.WriteLine($"Archivo Excel origen: {rutaArchivoExcel}");

        if (!File.Exists(rutaArchivoExcel))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n[ERROR] No se encontró el archivo Excel en: {rutaArchivoExcel}");
            Console.ResetColor();
            return;
        }

        try
        {
            var resultadoFinal = new Dictionary<string, HojaProcesamiento>();

            using (var stream = new FileStream(rutaArchivoExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
            {
                // Iterar sobre cada hoja del libro de Excel
                do
                {
                    string sheetName = reader.Name;
                    if (string.IsNullOrWhiteSpace(sheetName)) continue;

                    var registrosHoja = new List<DiaInfo>();
                    int filaActual = 1;

                    // Escanear todas las filas de la hoja desde el inicio (fila 1) hasta el final
                    while (reader.Read())
                    {
                        int colCount = reader.FieldCount;

                        // La fecha requerida está en la columna "B" (index 1)
                        object? rawFecha = colCount > 1 ? reader.GetValue(1) : null;
                        DateTime? fechaDate = ParsearFecha(rawFecha);

                        // Filtrar únicamente los días de Marzo y Abril de 2026
                        if (fechaDate.HasValue && fechaDate.Value.Year == 2026 && (fechaDate.Value.Month == 3 || fechaDate.Value.Month == 4))
                        {
                            string fechaFormatted = fechaDate.Value.ToString("d/MM/yyyy");

                            registrosHoja.Add(new DiaInfo
                            {
                                FechaHoja = fechaFormatted,
                                Fila = filaActual
                            });
                        }

                        filaActual++;
                    }

                    // Regla de deduplicación: "solo deja el primer registro fecha repetido"
                    var registrosFiltrados = registrosHoja
                        .GroupBy(r => r.FechaHoja)
                        .Select(g => g.First())
                        .OrderBy(r => r.Fila) // Mantener orden ascendente por fila
                        .ToList();

                    // Si la hoja tiene registros válidos para Marzo y Abril 2026, la agregamos al JSON
                    if (registrosFiltrados.Count > 0)
                    {
                        resultadoFinal[sheetName] = new HojaProcesamiento
                        {
                            CeldaLiquidos = "AA",
                            CeldaDescuento = "AA",
                            DiaProcesar = registrosFiltrados
                        };
                        Console.WriteLine($"[PROCESO] Hoja '{sheetName}': se encontraron {registrosFiltrados.Count} días únicos de Marzo/Abril 2026.");
                    }

                } while (reader.NextResult());
            }

            // 3. Serializar y guardar el nuevo archivo JSON
            var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
            string jsonResultado = JsonSerializer.Serialize(resultadoFinal, jsonOptions);

            File.WriteAllText(rutaSalidaJson, jsonResultado);

            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"\n[ÉXITO] Se ha generado con éxito el archivo JSON en:");
            Console.WriteLine(rutaSalidaJson);
            Console.ResetColor();

            Console.WriteLine($"Total de libros/hojas procesados con datos en Marzo/Abril 2026: {resultadoFinal.Count}");
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"\n[ERROR] Ocurrió un error inesperado al procesar el Excel: {ex.Message}");
            Console.ResetColor();
        }
    }

    #region Helpers de Conversión y Formato

    // Intenta parsear de manera robusta el objeto a DateTime
    private static DateTime? ParsearFecha(object? rawFecha)
    {
        if (rawFecha == null) return null;
        if (rawFecha is DateTime dt) return dt;
        
        if (DateTime.TryParse(rawFecha.ToString(), out DateTime parsedDate)) 
            return parsedDate;

        return null;
    }

    #endregion
}

public class HojaProcesamiento
{
    [JsonPropertyName("CeldaLiquidos")]
    public string CeldaLiquidos { get; set; } = "AA";

    [JsonPropertyName("CeldaDescuento")]
    public string CeldaDescuento { get; set; } = "AA";

    [JsonPropertyName("DiaProcesar")]
    public List<DiaInfo> DiaProcesar { get; set; } = new();
}

public class DiaInfo
{
    [JsonPropertyName("fecha_hoja")]
    public string? FechaHoja { get; set; }

    [JsonPropertyName("Fila")]
    public int Fila { get; set; }
}
