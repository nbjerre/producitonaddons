using System.Collections.Generic;

namespace WorksheetAPI.Models
{
    /// <summary>
    /// Requestmodel til opdatering af worksheet-linjer.
    /// </summary>
    public class WorksheetLinesUpdateDto
    {
        /// <summary>
        /// Liste af linjeobjekter fra UI'et, som mappes til SAP-felter i controlleren.
        /// </summary>
        public List<Dictionary<string, object>> Lines { get; set; } = new();
    }
}
