using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Octokit;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace GithubExcelExporter
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var githubOrganizationName = ConfigurationManager.AppSettings["github:organization-name"];
            var githubRepositoryName = ConfigurationManager.AppSettings["github:repository-name"];
            var githubUsername = ConfigurationManager.AppSettings["github:username"];
            var githubPassword = ConfigurationManager.AppSettings["github:password"];

            var basePath = ConfigurationManager.AppSettings["storage:basepath"];
            var fileName = ConfigurationManager.AppSettings["storage:filename"]; 

            var client = new GitHubClient(new ProductHeaderValue("issue-exporter"))
            {
                Credentials = new Credentials(githubUsername, githubPassword)
            };

            Console.WriteLine("INFO: Start exporting...");

            try
            {
                // Getting the complete workbook...
                var workbook = new HSSFWorkbook();

                // Getting the worksheet by its name...
                ISheet sheet = workbook.CreateSheet("Issues");

                int rowIndex = 0;
                IRow row;

                CreateHeader(sheet, rowIndex);
                
                rowIndex++;

                Console.WriteLine("INFO: Search issues from github repository " + githubRepositoryName);

                var currentPage = 1;

                ICollection<Issue> issues = null;

                do
                {
                    var result = Task.Run<IReadOnlyCollection<Issue>>(async () => await client.Issue.GetAllForRepository(githubOrganizationName, githubRepositoryName, new ApiOptions { StartPage = currentPage, PageSize = 30, PageCount = 1 }));
                    result.Wait();
                    issues = result.Result.Where(i => i.PullRequest == null).ToArray();

                    Console.WriteLine(string.Format("INFO: Collected {0} issues in page {1}.", issues.Count, currentPage));

                    FillFields(sheet, issues, ref rowIndex);

                    System.Threading.Thread.Sleep(1200); //// aguardar um tempo para o github não cancelar a requisição.
                    currentPage++;
                }
                while (issues.Count > 0); //// TODO: ver se há uma forma de capturar a quantidade de issues, assim não precisa ir checando página a página.

                var fileFullPath = Path.Combine(basePath, fileName);
                // Save the Excel spreadsheet to a file on the web server's file system
                using (var fileData = new FileStream(fileFullPath, System.IO.FileMode.Create))
                {
                    workbook.Write(fileData);
                }

                Console.WriteLine("INFO: Create excel with success in " + fileFullPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
            }

            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        private static void CreateHeader(ISheet sheet, int rowIndex)
        {
            var row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue("Número");
            row.CreateCell(1).SetCellValue("Título");
            row.CreateCell(2).SetCellValue("Descrição");
            row.CreateCell(3).SetCellValue("Labels");
            row.CreateCell(4).SetCellValue("Link");
            row.CreateCell(5).SetCellValue("Responsável");
        }

        private static void FillFields(ISheet sheet, ICollection<Issue> issues, ref int rowIndex)
        {
            foreach (Issue issue in issues)
            {
                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue(issue.Number);
                row.CreateCell(1).SetCellValue(issue.Title);
                row.CreateCell(2).SetCellValue(issue.Body);
                row.CreateCell(3).SetCellValue(string.Join(", ", issue.Labels.Select(l => l.Name)));
                row.CreateCell(4).SetCellValue(issue.Url.ToString());

                if (issue.Assignee != null)
                {
                    row.CreateCell(5).SetCellValue(issue.Assignee.Login ?? issue.Assignee.Name);
                }

                rowIndex++;
            }
        }
    }
}