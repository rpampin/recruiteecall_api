using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using RecruiteeCall.Models;
using RecruiteeCall.Services.RenderReport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace RecruiteeCall.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class RecruteeController : ControllerBase
    {
        protected readonly IConfiguration _configuration;
        protected readonly IRenderReport _renderReport;

        public RecruteeController(
            IConfiguration configuration,
            IRenderReport renderReport)
        {
            _configuration = configuration;
            _renderReport = renderReport;
        }

        string StripHTML(string input)
        {
            return Regex.Replace(input, "<.*?>", String.Empty);
        }

        [HttpGet]
        public async Task<IActionResult> Get([FromQuery] int candidateId = 21708756)
        {
            string baseUrl = $"https://api.recruitee.com/c/56989/interview/results?scope=all&candidate_id={candidateId}";
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _configuration.GetValue<string>("RecruteeToken"));

                using (HttpResponseMessage res = await client.GetAsync(baseUrl))
                {
                    using (HttpContent content = res.Content)
                    {
                        string jsonResponse = await content.ReadAsStringAsync();
                        if (jsonResponse != null)
                        {
                            int questionCount = 1;
                            string groupName = "";
                            JObject data = JObject.Parse(jsonResponse);

                            var adminId = data.SelectToken("interview_results[0].admin_id").Value<int>();
                            var updateAt = data.SelectToken("interview_results[0].updated_at").Value<DateTime>();
                            var templateName = data.SelectToken("interview_results[0].interview_template_name").Value<string>();
                            var candidateName = data.SelectToken("references").First(n => n["id"].Value<int>() == candidateId)["name"].Value<string>();
                            var AdminName = data.SelectToken("references").First(n => n["id"].Value<int>() == adminId)["first_name"].Value<string>();
                            var AdminLastName = data.SelectToken("references").First(n => n["id"].Value<int>() == adminId)["last_name"].Value<string>();
                            var questions = data.SelectToken("interview_results[0].interview_result_answers");

                            IList<Question> qs = new List<Question>();
                            qs.Add(new Question
                            {
                                Index = questionCount++,
                                Title = "Full name of consultant being evaluated",
                                Content = candidateName
                            });
                            qs.Add(new Question
                            {
                                Index = questionCount++,
                                Title = "Date of assessment",
                                Content = updateAt.ToString("MM/dd/yyyy")
                            });
                            qs.Add(new Question
                            {
                                Index = questionCount++,
                                Title = "Evaluator's name",
                                Content = AdminName + " " + AdminLastName
                            });

                            foreach (var q in questions)
                            {
                                var qTitle = q["question_title"] != null ? q["question_title"].ToString() : "";
                                var qContent = q["content"] != null ? q["content"].ToString() : "";

                                if (string.IsNullOrEmpty(qTitle))
                                {
                                    //groupName = StripHTML(q["question_hint"].ToString()) + ": ";
                                    continue;
                                }

                                qs.Add(new Question
                                {
                                    Index = questionCount++,
                                    Title = $"{groupName}{qTitle}",
                                    Content = qContent
                                });
                            }

                            var fileName = _renderReport.getReport(candidateId, candidateName, templateName, qs);
                            return File(new FileStream(fileName, FileMode.Open), $"application/{fileName.Split('.').Last()}", fileName.Split('\\').Last());
                        }

                        return NotFound();
                    }
                }
            }
        }
    }
}
