// ============================================================================
//  RefManagerLoader.cs — loads the primary library file (refmanager.json)
//  written by the Mirare desktop app:
//
//      { "articles": [ {Crossref-style article}, ... ] }
//
//  Field mapping (defensive — types vary between fetch platforms) lives in
//  ArticleMapper. Legacy/foreign shapes with an "entries" array are still
//  tolerated.
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using MirareCiteAddIn.Models;

namespace MirareCiteAddIn.Services
{
    public class RefManagerLoader
    {
        private readonly Logger _log;

        public RefManagerLoader(Logger log) { _log = log; }

        public List<Citation> Load(string jsonPath)
        {
            if (!File.Exists(jsonPath))
                throw new FileNotFoundException("refmanager.json not found", jsonPath);

            string json = File.ReadAllText(jsonPath);
            var root = JObject.Parse(json);

            if (root.GetValue("articles", StringComparison.OrdinalIgnoreCase) is JArray articles)
            {
                var citations = ArticleMapper.FromArticles(articles, CitationOrigin.Library);
                _log.Info($"Loaded library {Path.GetFileName(jsonPath)} ({citations.Count} articles)");
                return citations;
            }

            // Legacy shape: { "entries": [ ... ] }
            if (root.GetValue("entries", StringComparison.OrdinalIgnoreCase) is JArray entries)
            {
                var citations = ArticleMapper.FromArticles(entries, CitationOrigin.Library);
                _log.Info($"Loaded legacy library ({citations.Count} entries)");
                return citations;
            }

            throw new InvalidDataException(
                "refmanager.json: no 'articles' array found — is this a Mirare library file?");
        }
    }
}
