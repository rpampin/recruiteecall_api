using RecruiteeCall.Models;
using System.Collections.Generic;

namespace RecruiteeCall.Services.RenderReport
{
    public interface IRenderReport
    {
        string getReport(int candidateId, string candidateName, string templateName, IList<Question> questions);
    }
}
