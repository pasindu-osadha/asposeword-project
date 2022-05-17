namespace asposeword_project.Dtos.DocumentDtos
{
    public class RequestPDFDto
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public DateTime DateTaken { get; set; }
        public string CompletedBy { get; set; }
        public string Designation { get; set; }

        public List<DataTable> tableData { get; set; }
    }

    public class DataTable
    {
        public string Criteria { get; set; }
        public int Completed { get; set; }
        public string Findir { get; set; }
        public string Response { get; set; }
        public string Corrective { get; set; }
        public String  Action { get; set; }
    }
}
