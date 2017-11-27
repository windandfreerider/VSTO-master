using System.Collections.Generic;
using System.Linq;
using Layered.Application.Services.Dto;

namespace Layered.Application.Services
{
    /// <summary>
    /// 平均分计算服务
    /// </summary>
    public class AverageScoreService : IAverageScoreService
    {
        /// <summary>
        /// 计算
        /// </summary>
        public AverageScoreDto[] Calculation(List<AverageScoreDto> dtos)
        {
            return dtos.GroupBy(g => g.Subject).Select(s => new AverageScoreDto
            {
                Subject = s.Key,
                Score = s.Average(a => a.Score)
            }).ToArray();
        }
    }
}
