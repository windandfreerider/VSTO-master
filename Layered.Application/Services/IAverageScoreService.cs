using System.Collections.Generic;
using Layered.Application.Services.Dto;

namespace Layered.Application.Services
{
    /// <summary>
    /// 平均分计算服务
    /// </summary>
    public interface IAverageScoreService
    {
        /// <summary>
        /// 计算
        /// </summary>
        AverageScoreDto[] Calculation(List<AverageScoreDto> dtos);
    }
}