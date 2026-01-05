using OpenXmlPowerTools;
using OpenXmlPowerTools.DocumentBuilder;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xunit;

namespace OxPt
{
    public class DcTests
    {
        public class TestSource
        {
            public string TemplateFile { get; set; }
            public string DataFile { get; set; }
            public string DocumentFile { get; set; }
            public int Start { get; set; }
            public int Count { get; set; }
            public bool KeepSections { get; set; }
            public bool DiscardHeadersAndFootersInKeptSections { get; set; }
            public string InsertId { get; set; }

            public Source ToSource(DirectoryInfo sourceDir)
            {
                DocxSource result;
                if (DocumentFile != null)
                {
                    var documentDocx = new FileInfo(Path.Combine(sourceDir.FullName, DocumentFile));
                    result = new DocxSource(documentDocx.FullName);
                }
                else
                {
                    var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, TemplateFile));
                    var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, DataFile));
                    XElement xmldata = XElement.Load(dataFile.FullName);
                    result = new TemplateSource(templateDocx.FullName, xmldata);
                }
                if (InsertId != result.InsertId) result.InsertId = InsertId;
                if (Start != result.Start) result.Start = Start;
                if (Count > 0 && Count != result.Count) result.Count = Count;
                if (KeepSections != result.KeepSections) result.KeepSections = KeepSections;
                if (DiscardHeadersAndFootersInKeptSections != result.DiscardHeadersAndFootersInKeptSections) result.DiscardHeadersAndFootersInKeptSections = DiscardHeadersAndFootersInKeptSections;
                return result;
            }
        }

        private static void ValidateDocx(string filename)
        {
            ValidateDocx(WordprocessingDocument.Open(filename, true));
        }

        private static void ValidateDocx(WmlDocument result)
        {
            using (var ms = new MemoryStream())
            {
                ms.Write(result.DocumentByteArray, 0, result.DocumentByteArray.Length);
                ValidateDocx(WordprocessingDocument.Open(ms, true));
            }
        }

        private static void ValidateDocx(WordprocessingDocument doc)
        {
            using (doc)
            {
                var v = new OpenXmlValidator();
                var valErrors = v.Validate(doc).Where(ve => !s_BuilderExpectedErrors.Contains(ve.Description));

#if false
                StringBuilder sb = new StringBuilder();
                foreach (var item in valErrors.Select(r => r.Description).OrderBy(t => t).Distinct())
	            {
		            sb.Append(item).Append(Environment.NewLine);
	            }
                string z = sb.ToString();
                Console.WriteLine(z);
#endif

                Assert.Empty(valErrors);
            }
        }

        private static async Task<WmlDocument> DoConcatTest(TestSource[] testSources, string destinationName)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/DC/");
            var sources = testSources.Select(src => src.ToSource(sourceDir)).ToList();
            var result = await DocumentComposer.ComposeDocument(sources);
            var composedDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, destinationName));
            result.SaveAs(composedDocx.FullName);
            ValidateDocx(result);
            return result;
        }

        private static async Task<WmlDocument> DoInsertTest(string template, string data, TestSource[] testSources, string destinationName)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/DC/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, template));
            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            XElement xmlData = null;
            if (!string.IsNullOrWhiteSpace(data))
            {
                var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, data));
                xmlData = XElement.Load(dataFile.FullName);
            }
            List<Source> sources = null;
            if (testSources != null)
            {
                sources = testSources.Select(src => src.ToSource(sourceDir)).ToList();
            }
            var result = await DocumentComposer.ComposeDocument(wmlTemplate, xmlData, sources);
            var composedDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, destinationName));
            result.SaveAs(composedDocx.FullName);
            ValidateDocx(result);
            return result;
        }

        [Fact]
        public async void DC001_DocumentComposerConcat()
        {
            var sourceArray = new TestSource[]
            {
                new TestSource()
                {
                    TemplateFile = "DC001-Template1.docx",
                    DataFile = "DC001-Data1.xml",
                    KeepSections = true,
                },
                new TestSource()
                {
                    TemplateFile = "DC001-Template2.docx",
                    DataFile = "DC001-Data2.xml",
                    KeepSections = true,
                },
            };
            await DoConcatTest(sourceArray, "DC001_DocumentComposerConcat_composed.docx");
        }

        [Fact]
        public async void DC002_DocumentComposerInsertStatic()
        {
            string template = "DC-MainSimpleInsId.docx"; // has direct insert with Id="1"
            string data = "DC-CustomerCheryl.xml";
            var sourceArray = new TestSource[]
            {
                new TestSource()
                {
                    TemplateFile = "DC-SimpleBulletRed.docx",
                    DataFile = data,
                    InsertId = "1", // this identifier is in an <Insert> element in the template
                },
            };
            await DoInsertTest(template, data, sourceArray, "DC002_DocumentComposerInsertStatic_composed.docx");
        }

        [Fact]
        public async void DC003_DocumentComposerInsertAuto()
        {
            string template = "DC-MainSimpleInsAuto.docx"; // has implicit insert of DC-SimpleBulletRed.docx
            string data = "DC-CustomerCheryl.xml";
            await DoInsertTest(template, data, null, "DC003_DocumentComposerInsertAuto_composed.docx");
        }

        [Fact]
        public async void DC004_DocumentComposerInsertAutoStatic()
        {
            string template = "DC-Main2SectInsAuto.docx"; // has implicit insert of DC-StaticInsertedWideMargin.docx
            string data = null;
            await DoInsertTest(template, data, null, "DC004_DocumentComposerInsertAutoStatic_composed.docx");
        }

        [Fact]
        public async void DC005_DocumentComposerInsertKeepSectNoBreaks()
        {
            string template = "DC-Main1SectInsId.docx"; // has direct insert with Id="1"
            string data = null;
            var sourceArray = new TestSource[]
            {
                new TestSource()
                {
                    DocumentFile = "DC-StaticInsertedWideMargin.docx",
                    InsertId = "1", // this identifier is in an <Insert> element in the template
                    KeepSections = true,
                },
            };
            await DoInsertTest(template, data, sourceArray, "DC005_DocumentComposerInsertKeepSectNoBreaks_composed.docx");
            // todo: now assert that...
            //   - the resulting document has 2 sections
            //   - the 1st section has the headers/footers/margins from the inserted doc (DC005-Inserted.docx)
            //   - the 2nd section has the headers/footers/margins from the parent (DC005-Insert.docx)
        }

        [Fact]
        public async void DC006_DocumentComposerInsertKeepSectWithBreakBefore()
        {
            string template = "DC-Main2SectInsId.docx"; // has direct insert with Id="1"
            string data = null;
            var sourceArray = new TestSource[]
            {
                new TestSource()
                {
                    DocumentFile = "DC-StaticInsertedWideMargin.docx",
                    InsertId = "1", // this identifier is in an <Insert> element in the template
                    KeepSections = true,
                },
            };
            await DoInsertTest(template, data, sourceArray, "DC006_DocumentComposerInsertKeepSectWithBreakBefore_composed.docx");
            // todo: now assert that...
            //   - the resulting document has 3 sections
            //   - the 1st and last sections still have the same properties as they did in DC-Main2SectInsId.docx
            //   - the 2nd section has the headers/footers/margins from inserted DC-StaticInsertedWideMargin.docx
        }

        [Fact]
        public async void DC008_DocumentComposerInsertKeepSectWithBreaksBeforeAndAfter()
        {
            string template = "DC-Main2SectInsId.docx"; // has direct insert with Id="1"
            string data = null;
            var sourceArray = new TestSource[]
            {
                new TestSource()
                {
                    TemplateFile = "DC-MarginConditional.docx",
                    InsertId = "1", // this identifier is in an <Insert> element in the template
                    DataFile = "DC-MarginConditionalWide.xml",
                    KeepSections = true,
                },
            };
            await DoInsertTest(template, data, sourceArray, "DC008_DocumentComposerInsertKeepSectWithBreaksBeforeAndAfter_composed.docx");
            // todo: now assert that...
            //   - the resulting document has 3 sections
            //   - the 1st and last sections still have the same properties as they did in DC-Main2SectInsId.docx
            //   - the 2nd section has the headers/footers/WIDE margins from DC-MarginConditional.docx
        }

        [Fact]
        public async void DC009_DocumentComposerInsertKeepSectWithoutConditionalBreakAfter()
        {
            string template = "DC-Main2SectInsId.docx"; // has direct insert with Id="1"
            string data = null;
            var sourceArray = new TestSource[]
            {
                new TestSource()
                {
                    TemplateFile = "DC-MarginConditional.docx",
                    InsertId = "1", // this identifier is in an <Insert> element in the template
                    DataFile = "DC-MarginConditionalNotWide.xml",
                    KeepSections = true,
                },
            };
            await DoInsertTest(template, data, sourceArray, "DC009_DocumentComposerInsertKeepSectWithoutConditionalBreakAfter_composed.docx");
            // todo: now assert that...
            //   - the resulting document has 3 sections
            //   - the 1st and last sections still have the same properties as they did in DC-Main2SectInsId.docx
            //   - the 2nd section has the headers/footers/REGULAR margins from DC-MarginConditional.docx
        }

        [Fact]
        public async void DC010_DocumentComposerInsertIndirect()
        {
            string template = "DC-MainInsertIndirect.docx"; // has no insert, but has Content control selecting data "./Indirect"
            string data = "DC-CustomerJuliaIndirect.xml"; // has data point 'Indirect' => "oxpt://DocumentAssembler/insert/1"
            var sourceArray = new TestSource[]
            {
                new TestSource()
                {
                    TemplateFile = "DC-SimpleBulletRed.docx",
                    InsertId = "1", // this identifier was in a URI in the XML data above
                    DataFile = "DC-CustomerCheryl.xml", // can be the same as, or different from, data used in parent assemblies
                },
            };
            var afterAssembling = await DoInsertTest(template, data, sourceArray, "DC010_DocumentComposerInsertIndirect_composed.docx");

            // assert that composition happened as expected...
            var mainPart = afterAssembling.MainDocumentPart;
            var paragraphs = mainPart.Element(W.body).Elements(W.p);
            string para1Text = paragraphs.First().Value;
            string para2Text = paragraphs.ElementAt(1).Value;
            Assert.Equal("This is a parent document, Julia.", para1Text);
            Assert.Equal("Cheryl's bulleted item!", para2Text);
        }

        private static readonly List<string> s_BuilderExpectedErrors = new List<string>()
        {
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowFirstColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowLastColumn' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddHBand' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddVBand' attribute is not declared.",
            "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:updateFields'.",
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:name' has invalid value 'useWord2013TrackBottomHyphenation'. The Enumeration constraint failed.",
            "The 'http://schemas.microsoft.com/office/word/2012/wordml:restartNumberingAfterBreak' attribute is not declared.",
            "Attribute 'id' should have unique value. Its current value '",
            "The 'urn:schemas-microsoft-com:mac:vml:blur' attribute is not declared.",
            "Attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:id' should have unique value. Its current value '",
            "The element has unexpected child element 'http://schemas.microsoft.com/office/word/2012/wordml:",
            "The element has invalid child element 'http://schemas.microsoft.com/office/word/2012/wordml:",
            "The 'urn:schemas-microsoft-com:mac:vml:complextextbox' attribute is not declared.",
            "http://schemas.microsoft.com/office/word/2010/wordml:",
            "http://schemas.microsoft.com/office/word/2008/9/12/wordml:",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:allStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:customStyles' attribute is not declared.",
            "The element has invalid child element 'http://schemas.microsoft.com/office/word/2010/wordml:ligatures'.",
        };
    }
}
