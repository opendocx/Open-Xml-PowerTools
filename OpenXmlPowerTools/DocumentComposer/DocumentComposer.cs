using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    public static class DocumentComposer
    {
        public static async Task<WmlDocument> ComposeDocument(WmlDocument templateDoc, XmlDocument data)
        {
            return await ComposeDocument(templateDoc, data, null);
        }

        public static async Task<WmlDocument> ComposeDocument(WmlDocument templateDoc, XmlDocument data, List<DocumentBuilder.Source> insertSources)
        {
            XDocument xDoc = data.GetXDocument();
            return await ComposeDocument(templateDoc, xDoc.Root, insertSources);
        }

        // Compose document via AUTO insertion (insertion sources inferred from template)
        public static async Task<WmlDocument> ComposeDocument(WmlDocument templateDoc, XElement data)
        {
            return await ComposeDocument(templateDoc, data, null);
        }

        // Compose document via insertion (a la Document Builder 2.0 WITH DocumentAssembler)
        public static async Task<WmlDocument> ComposeDocument(WmlDocument templateDoc, XElement data, List<DocumentBuilder.Source> insertSources)
        {
            await Task.Yield();
            var sourcesProvided = insertSources != null && insertSources.Count > 0;
            // assemble document
            var mainDoc = DocumentAssembler.AssembleDocument(templateDoc, data, out DocumentAssembler.AssembleResult results);
            var errors = new List<string>();
            if (results.HasError)
            {
                errors.Add("Template Error");
            }
            if (results.Inserts.Any())
            {
                var sources = new List<DocumentBuilder.Source>
                {
                    new DocxSource(mainDoc, true)
                };
                if (sourcesProvided)
                {
                    // verify that all inserts encountered in the template are among those provided, otherwise err
                }
                else
                {
                    // no sources provided -- assume AUTO inserts (infer sources from what was in the template)
                    var dir = Path.GetDirectoryName(templateDoc.FileName);
                    insertSources = results.Inserts.Select(asmInsert => GetSourceFromInsertResult(asmInsert, dir)).ToList();
                    if (insertSources.Any(s => s == null))
                    {
                        insertSources = null;
                        errors.Add("Insert source not found");
                    }
                    else
                    {
                        sourcesProvided = true;
                    }
                }
                if (sourcesProvided)
                {
                    var tasks = insertSources.Select(source =>
                    {
                        var templateSource = source as TemplateSource;
                        if (templateSource != null && templateSource.WmlDocument == null)
                        {
                            if (!templateSource.HasError)
                                return templateSource.DoAssembly();
                            else
                                errors.Add("Insert Error");
                        }
                        return Task.CompletedTask;
                    });
                    await Task.WhenAll(tasks);
                }
                if (errors.Count == 0)
                {
                    sources.AddRange(insertSources);
                    var result = DocumentBuilder.DocumentBuilder.BuildDocument(sources);
                    return result;
                }
                else
                {
                    return null;
                }
            }
            else
            { // no inserts
                return mainDoc;
            }
        }

        // Compose document via concatenation (a la DocumentBuilder 1.0)
        public static async Task<WmlDocument> ComposeDocument(List<DocumentBuilder.Source> sources)
        {
            var errors = new List<string>();
            var tasks = sources.Select(source =>
            {
                var templateSource = source as TemplateSource;
                if (templateSource != null && templateSource.WmlDocument == null)
                {
                    if (!templateSource.HasError)
                        return templateSource.DoAssembly();
                    else
                        errors.Add("Source Error");
                }
                return Task.CompletedTask;
            });
            await Task.WhenAll(tasks);
            if (errors.Count == 0)
            {
                return DocumentBuilder.DocumentBuilder.BuildDocument(sources);
            }
            else
            {
                return null;
            }
        }

        private static DocumentBuilder.Source GetSourceFromInsertResult(DocumentAssembler.AssembleInsert insert, string dirName)
        {
            var filename = GetFilenameFromInsertId(insert.Id, dirName, insert.Data != null);
            if (filename != null && File.Exists(filename))
            {
                if (insert.Data == null)
                {
                    return new DocxSource(filename, insert.Id);
                }
                else
                {
                    return new TemplateSource(filename, insert.Data, insert.Id);
                }
            }
            return null;
        }

        private static string GetFilenameFromInsertId(string insertId, string dirName, bool hasData)
        {
            if (string.IsNullOrWhiteSpace(insertId))
                return null;
            // ensure insertId is not trying to do anything funny
            if (insertId.Contains('/') || insertId.Contains('\\') || insertId.Contains(".."))
                return null;
            if (hasData)
            {
                Match match = s_endDigits.Match(insertId);
                if (match.Success)
                    insertId = insertId[..match.Index];
            }
            return Path.Combine(dirName, insertId);
        }

        private static readonly Regex s_endDigits = new Regex("@\\d+$");
    }
}
