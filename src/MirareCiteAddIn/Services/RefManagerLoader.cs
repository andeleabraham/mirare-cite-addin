// ============================================================================
//  RefManagerLoader.cs — loads a plain .refmanager.json (the primary
//  library file).  Unlike .mrrcite (a ZIP), this is just JSON — read it
//  straight off disk.
//
//  See docs/data-formats.md for the schema.  The structure mirrors what
//  the RefManagerLibrary model expects:
//    {
//      "libraryId": "...",
//      "name": "My Library",
//      "entries": [
//        { "id": "...", "title": "...", "authors": [...], ... },
//        ...
//      ]
//    }
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
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
                throw new FileNotFoundException(".refmanager.json not found", jsonPath);

            string json = File.ReadAllText(jsonPath, Encoding.UTF8);
            var lib = System.Text.Json.JsonSerializer.Deserialize<RefManagerLibrary>(json);

            _log.Info($"Loaded library {lib.Name} ({lib.Entries?.Count ?? 0} entries)");

            var citations = new List<Citation>();
            if (lib.Entries == null) return citations;
            foreach (var e in lib.Entries)
            {
                citations.Add(new Citation
                {
                    Id = e.Id,
                    Title = e.Title,
                    Authors = e.Authors,
                    Year = e.Year,
                    Journal = e.Journal,
                    Doi = e.Doi,
                    Url = e.Url,
                    MiRNA = e.MiRNA,
                    TargetGene = e.TargetGene,
                    EvidenceType = e.EvidenceType,
                    SourceDb = e.SourceDb,
                    Origin = CitationOrigin.Library
                });
            }
            return citations;
        }
    }
}
