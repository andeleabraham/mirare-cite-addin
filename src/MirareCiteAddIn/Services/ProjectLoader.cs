// ============================================================================
//  ProjectLoader.cs — loads a .mrrcite project file (a ZIP) written by the
//  Mirare desktop app. We never write to it — read-only.
//
//  Layout (mirrors what the app exports):
//      <archive>.mrrcite  (ZIP)
//        └── project.json
//            { "<ProjectName>": { "articles": [ {...}, ... ] } }
//
//  The articles use the same Crossref-style shape as refmanager.json, so
//  the mapping lives in ArticleMapper. For robustness we also accept:
//    * project.json = { "articles": [...] }  (no project-name wrapper)
//    * an embedded refs.refmanager.json alongside project.json
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Newtonsoft.Json.Linq;
using MirareCiteAddIn.Models;

namespace MirareCiteAddIn.Services
{
    public class ProjectLoader
    {
        private readonly Logger _log;

        public ProjectLoader(Logger log) { _log = log; }

        /// <summary>
        /// Loads a .mrrcite file and returns the citation list.
        /// Throws on malformed archives — the picker dialog catches.
        /// </summary>
        public List<Citation> Load(string mrrcitePath)
        {
            if (!File.Exists(mrrcitePath))
                throw new FileNotFoundException(".mrrcite file not found", mrrcitePath);

            using (var zip = ZipFile.OpenRead(mrrcitePath))
            {
                var projectEntry = zip.GetEntry("project.json")
                                  ?? throw new InvalidDataException(
                                      ".mrrcite archive is missing project.json — not a valid Mirare project.");

                var citations = new List<Citation>();
                JObject root;
                using (var s = projectEntry.Open())
                using (var reader = new StreamReader(s))
                {
                    root = JObject.Parse(reader.ReadToEnd());
                }

                // Shape A: { "articles": [...] }
                if (root.GetValue("articles", StringComparison.OrdinalIgnoreCase) is JArray direct)
                {
                    citations.AddRange(ArticleMapper.FromArticles(direct, CitationOrigin.Project));
                }
                else
                {
                    // Shape B: { "<ProjectName>": { "articles": [...] } }
                    // — the app wraps everything under the project name.
                    foreach (var prop in root.Properties())
                    {
                        if (!(prop.Value is JObject section)) continue;
                        if (section.GetValue("articles", StringComparison.OrdinalIgnoreCase)
                                is JArray arr)
                        {
                            citations.AddRange(ArticleMapper.FromArticles(arr, CitationOrigin.Project));
                        }
                    }
                }

                // Optional embedded library — append entries not already present.
                var refsEntry = zip.GetEntry("refs.refmanager.json");
                if (refsEntry != null)
                {
                    var seen = new HashSet<string>(citations.Select(c => c.Id));
                    JObject refsRoot;
                    using (var s = refsEntry.Open())
                    using (var reader = new StreamReader(s))
                    {
                        refsRoot = JObject.Parse(reader.ReadToEnd());
                    }
                    if (refsRoot.GetValue("articles", StringComparison.OrdinalIgnoreCase)
                            is JArray arr)
                    {
                        foreach (var c in ArticleMapper.FromArticles(arr, CitationOrigin.Project))
                            if (seen.Add(c.Id)) citations.Add(c);
                    }
                }

                _log.Info($"Loaded project {Path.GetFileName(mrrcitePath)} ({citations.Count} citations)");
                return citations;
            }
        }
    }
}
