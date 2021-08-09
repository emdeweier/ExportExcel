using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using LogActivitty.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Nest;

namespace LogActivitty.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LogsController : ControllerBase
    {
        ElasticClient _elasticClient = null;
        private readonly ILogger<LogsController> _logger;
        public LogsController(ILogger<LogsController> logger)
        {
            var _settings = new ConnectionSettings(new Uri("http://localhost:9200")).DefaultIndex("products");
            _elasticClient = new ElasticClient(_settings);
            _logger = logger;
        }

        [HttpGet("result/{id}")]
        public List<Product> GetResult(string id)
        {
            _logger.LogInformation("called get product");
            var response = _elasticClient.Search<Product>(s => s
                    .From(0)
                    .Take(10)
                    .Query(qry => qry
                        .Bool(b => b
                            .Must(m => m
                                .QueryString(qs => qs
                                    .DefaultField("id")
                                    .Query(id))))));
            return response.Documents.ToList();
        }

        public void AddNewIndex(Product model)
        {
            _elasticClient.IndexAsync<Product>(model, null);
        }

        [HttpPost("post")]
        public void AddNewIndexTest(Product model)
        {
            _logger.LogInformation("called post product");

            _elasticClient.IndexAsync<Product>(model, null);
            AddNewIndex(new Product
            {
                Id = 3,
                Name = "Celana",
                Description = "Celana",
                Brand = "Gucci",
                Price = "9.879 million"
            });

        }
    }
}