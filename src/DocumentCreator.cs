using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace DocxReview;

/// <summary>
/// Creates new documents from the bundled NIH template.
/// Optionally populates the new document using an EditManifest.
/// </summary>
public class DocumentCreator
{
    private const string EmbeddedTemplateName = "DocxReview.templates.nih-standard.docx";

    /// <summary>
    /// Create a new document from the template, optionally populating it.
    /// </summary>
    public CreateResult Create(
        string outputPath,
        EditManifest? manifest,
        string author,
        string? templatePath,
        bool dryRun)
    {
        string templateLabel = templatePath ?? "nih-standard (embedded)";

        if (templatePath != null && !File.Exists(templatePath))
            throw new FileNotFoundException($"Template not found: {templatePath}");

        var result = new CreateResult
        {
            Template = templateLabel,
            Output = dryRun ? null : outputPath,
            Populated = manifest != null,
            Success = true
        };

        // For dry-run without manifest, nothing to do — just report success
        if (dryRun && manifest == null)
            return result;

        // Extract template to a temp file. DocumentEditor.Process() will copy
        // from this temp file to the output path (avoids same-file copy issue).
        string tempTemplate = Path.Combine(Path.GetTempPath(), $"docx-review-template-{Guid.NewGuid()}.docx");
        CopyTemplate(templatePath, tempTemplate);

        try
        {
            if (manifest != null)
            {
                // Let DocumentEditor.Process() handle the copy from temp → output and apply edits
                var editor = new DocumentEditor(author);
                var editResult = editor.Process(tempTemplate, dryRun ? "" : outputPath, manifest, dryRun);

                result.ChangesAttempted = editResult.ChangesAttempted;
                result.ChangesSucceeded = editResult.ChangesSucceeded;
                result.CommentsAttempted = editResult.CommentsAttempted;
                result.CommentsSucceeded = editResult.CommentsSucceeded;
                result.Results = editResult.Results;
                result.Success = editResult.Success;
            }
            else
            {
                // No manifest — just copy the template to the output path
                File.Copy(tempTemplate, outputPath, true);
            }
        }
        finally
        {
            if (File.Exists(tempTemplate))
                File.Delete(tempTemplate);
        }

        return result;
    }

    private static void CopyTemplate(string? templatePath, string destinationPath)
    {
        var dir = Path.GetDirectoryName(destinationPath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir);

        if (templatePath != null)
        {
            File.Copy(templatePath, destinationPath, true);
        }
        else
        {
            ExtractEmbeddedTemplate(destinationPath);
        }
    }

    private static void ExtractEmbeddedTemplate(string destinationPath)
    {
        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream(EmbeddedTemplateName)
            ?? throw new InvalidOperationException(
                $"Embedded template '{EmbeddedTemplateName}' not found in assembly.");

        using var fileStream = File.Create(destinationPath);
        stream.CopyTo(fileStream);
    }
}
