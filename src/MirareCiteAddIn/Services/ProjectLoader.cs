// ============================================================================
//  ProjectLoader.cs — loads a .mrrcite file (which is a ZIP) and extracts
//  project.json + refs.refmanager.json (if present) into Citation objects.
//
//  The .mrrcite format is documented in docs/data-formats.md:
//    * It is a standard ZIP archive (so System.IO.Compression handles it).
//    * Inside:  project.json (primary), refs.refmanager.json (optional),
//               notes/, attachments/, etc. (ignored by the add-in).
//    * We never write to the .mrrcite file — we only read it.
// ============================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
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
                                      ".mrrcite archive is missing project.json — not a valid Mirare Cite project.");
                var project = ReadJson<MirareProject>(projectEntry);
                _log.Info($"Loaded project {project.Name} ({project.Citations?.Count ?? 0} citations)");

                var citations = new List<Citation>();
                if (project.Citations != null)
                {
                    foreach (var e in project.Citations)
                    {
                        citations.Add(ToCitation(e, CitationOrigin.Project));
                    }
                }

                // Optional embedded library — append its entries too, but
                // dedupe by Id (project entries take priority since the user
                // may have annotated them inside the project).
                var refsEntry = zip.GetEntry("refs.refmanager.json");
                if (refsEntry != null)
                {
                    var lib = ReadJson<RefManagerLibrary>(refsEntry);
                    var seen = new HashSet<string>(citations.ConvertAll(c => c.Id));
                    foreach (var e in lib.Entries)
                    {
                        if (seen.Contains(e.Id)) continue;
                        citations.Add(ToCitation(e, CitationOrigin.Project));
                        seen.Add(e.Id);
                    }
                }
                return citations;
            }
        }

        // ─────────────────────────────────────────────────────────────────
        //  Mapping helpers — project entry / library entry → Citation.
        //  Kept explicit (not AutoMapper) so the field set is visible.
        // ─────────────────────────────────────────────────────────────────
        private static Citation ToCitation(ProjectCitationEntry e, CitationOrigin origin)
            => new Citation
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
                Origin = origin
            };

        private static Citation ToCitation(RefManagerEntry e, CitationOrigin origin)
            => new Citation
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
                Origin = origin
            };

        // ─────────────────────────────────────────────────────────────────
        //  ZIP entry → JSON deserialization.  We assume UTF-8 (the .mrrcite
        //  spec mandates it) but tolerate a BOM.
        // ─────────────────────────────────────────────────────────────────
        private static T ReadJson<T>(ZipArchiveEntry entry)
        {
            using (var s = entry.Open())
            using (var ms = new MemoryStream())
            {
                s.CopyTo(ms);
                // Strip BOM if present.
                var bytes = ms.ToArray();
                if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                    bytes = bytes.Skip(3).ToArray();
                string json = Encoding.UTF8.GetString(bytes);
                return System.Text.Json.JsonSerializer.Deserialize<T>(json);
            }
        }
    }
}
