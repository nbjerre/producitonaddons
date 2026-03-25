namespace WebAPI.Models
{
    /// <summary>
    /// Request til annullering, frigivelse, færdigmelding eller print af produktionsordrer relateret til en salgsordre.
    /// </summary>
    public class CancelAllProductionsRequest
    {
        /// <summary>
        /// SAP DocEntry for salgsordren.
        /// </summary>
        public int SalesOrderDocEntry { get; set; }

        /// <summary>
        /// Alternativ SAP DocNum hvis DocEntry ikke kendes.
        /// </summary>
        public int? SalesOrderDocNum { get; set; }

        /// <summary>
        /// Valgfrit id for et tidligere genereret printdokument.
        /// </summary>
        public string? DocumentId { get; set; }
    }
}
