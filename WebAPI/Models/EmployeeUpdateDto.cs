namespace WorksheetAPI.Models
{
    public class EmployeeUpdateDto
    {
        public string? FirstName { get; set; }
        public string? LastName { get; set; }
        public string? JobTitle { get; set; }
        public string? Position { get; set; }
        public string? Department { get; set; }
        public string? Branch { get; set; }
        public string? Manager { get; set; }
        public string? CostCenterCode { get; set; }
        public string? OfficePhone { get; set; }
        public string? MobilePhone { get; set; }
        public string? eMail { get; set; }
        public string? StartDate { get; set; }
        public string? TerminationDate { get; set; }
        public decimal? Salary { get; set; }
        public string? BankAccount { get; set; }
        public string? PassportNumber { get; set; }
        public string? MartialStatus { get; set; }
        public string? Active { get; set; }
        // Custom fields
        public string? U_AnsatNr { get; set; }
        public string? U_AIG_IK { get; set; }
        public string? U_AIG_LOEN { get; set; }
        public string? U_AIG_KCIF { get; set; }
        public string? U_AIG_FLX { get; set; }
        public string? U_AIG_JBTD { get; set; }
        public string? U_AIG_GPER { get; set; }
        public string? U_AIG_ARBP { get; set; }
        public string? U_AIG_PNO { get; set; }
        public string? U_AIG_SKF { get; set; }
        public string? U_AIG_FSTL { get; set; }
        public string? U_AIG_FUNK { get; set; }
        public string? U_AIG_AFLE { get; set; }
        public string? U_AIG_AOT { get; set; }
        public string? U_AIG_KLAN { get; set; }
        public string? U_AIG_VP { get; set; }
        public string? U_AIG_JBTDQR { get; set; }
        public int? U_RCS_TTP { get; set; }
        public string? U_RCS_LL { get; set; }
        public string? U_RCS_FFD { get; set; }
        public string? U_RCS_FJ { get; set; }
        public int? U_RCS_SHP { get; set; }
    }
}
