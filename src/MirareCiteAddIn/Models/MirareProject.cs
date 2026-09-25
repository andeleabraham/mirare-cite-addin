// ============================================================================
//  MirareProject.cs — in-memory model for a .mrrcite project file.
//
//  On disk, a .mrrcite file is a ZIP archive whose contents include:
//    project.json         ← primary citation list + project metadata
//    notes/...            ← user notes (irrelevant for citation)
//    attachments/...      ← PDFs, etc. (irrelevant for citation)
//    refs.refmanager.json ← optional embedded copy of the source library
//
//  We only care about project.json (and refs.refmanager.json if present).
//  See docs/data-formats.md for the full schema.
// ============================================================================

using System.Collections.Generic;

namespace MirareCiteAddIn.Models
{
    public class MirareProject
    {
        public string ProjectId { get; set; }
        public string Name { get; set; }
        public string CreatedAt { get; set; }
        public string LastModified { get; set; }
        public List<ProjectCitationEntry> Citations { get; set; } = new List<ProjectCitationEntry>();
        public ProjectMetadata Metadata { get; set; }
    }

    public class ProjectCitationEntry
    {
        // Same id semantics as Citation.Id — used to dedupe across library/project/remote.
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

        // Project-specific: annotations the user added inside the project.
        public string UserNote { get; set; }
        public List<string> Tags { get; set; } = new List<string>();
    }

    public class ProjectMetadata
    {
        public string Owner { get; set; }
        public string Description { get; set; }
    }
}
