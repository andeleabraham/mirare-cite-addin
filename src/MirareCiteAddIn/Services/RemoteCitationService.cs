// ============================================================================
//  RemoteCitationService.cs — queries the remote Mirare API over HTTPS and
//  parses the JSON response into Citation objects.
//
//  The endpoint URL is configurable in Settings (default:
//  https://api.mirare.example.org/cite).  The expected response shape is:
//
//    {
//      "results": [
//        { "id": "...", "title": "...", "authors": [...], "year": 2022,
//          "journal": "...", "doi": "...", "url": "...",
//          "miRNA": "hsa-miR-21-5p", "targetGene": "PTEN",
//          "evidenceType": "validated", "sourceDb": "miRTarBase" },
//        ...
//      ]
//    }
//
//  This service is intentionally simple — a single GET with a query string.
//  If your real endpoint uses POST / auth headers / paging, extend the
//  QueryAsync method below.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using MirareCiteAddIn.Models;

namespace MirareCiteAddIn.Services
{
    public class RemoteCitationService
    {
        private readonly Logger _log;
        private readonly string _endpoint;
        // Reuse a single HttpClient across calls — best practice on .NET
        // Framework 4.8 (avoids socket exhaustion).
        private static readonly HttpClient _http;

        static RemoteCitationService()
        {
            _http = new HttpClient(new HttpClientHandler
            {
                // Trust the system trust store; if your server uses a CA
                // not in the Windows trust store, install it there rather
                // than disabling validation.
            });
            _http.Timeout = TimeSpan.FromSeconds(15);
            // Identify ourselves to the API.
            _http.DefaultRequestHeaders.UserAgent.ParseAdd(
                "MirareCite-WordAddin/1.0");
        }

        public RemoteCitationService(string endpoint, Logger log)
        {
            _endpoint = endpoint;
            _log = log;
        }

        public async Task<List<Citation>> QueryAsync(string searchTerm)
        {
            if (string.IsNullOrWhiteSpace(_endpoint))
                throw new InvalidOperationException("Remote endpoint is not configured.");

            string url = $"{_endpoint}?q={Uri.EscapeDataString(searchTerm ?? "")}";
            _log.Info($"RemoteQuery: GET {url}");
            var resp = await _http.GetAsync(url);
            resp.EnsureSuccessStatusCode();
            string body = await resp.Content.ReadAsStringAsync();

            var payload = JsonSerializer.Deserialize<RemoteResponse>(body);
            var citations = new List<Citation>();
            if (payload?.Results == null) return citations;
            foreach (var r in payload.Results)
            {
                citations.Add(new Citation
                {
                    Id = r.Id,
                    Title = r.Title,
                    Authors = r.Authors,
                    Year = r.Year,
                    Journal = r.Journal,
                    Doi = r.Doi,
                    Url = r.Url,
                    MiRNA = r.MiRNA,
                    TargetGene = r.TargetGene,
                    EvidenceType = r.EvidenceType,
                    SourceDb = r.SourceDb,
                    Origin = CitationOrigin.Remote
                });
            }
            _log.Info($"RemoteQuery: returned {citations.Count} results");
            return citations;
        }

        // Local DTO mirroring the expected JSON response.
        private class RemoteResponse
        {
            public List<RemoteEntry> Results { get; set; }
        }
        private class RemoteEntry
        {
            public string Id { get; set; }
            public string Title { get; set; }
            public List<string> Authors { get; set; } = new List<string>();
            public int? Year { get; set; }
            public string Journal { get; set; }
            public string Doi { get; set; }
            public string Url { get; set; }
            public string MiRNA { get; set; }
            public string TargetGene { get; set; }
            public string EvidenceType { get; set; }
            public string SourceDb { get; set; }
        }
    }
}
