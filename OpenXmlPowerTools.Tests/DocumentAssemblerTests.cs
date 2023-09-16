// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class DaTests
    {
        [Theory]
        [InlineData("DA001-TemplateDocument.docx", "DA-Data.xml", false)]
        [InlineData("DA002-TemplateDocument.docx", "DA-DataNotHighValueCust.xml", false)]
        [InlineData("DA003-Select-XPathFindsNoData.docx", "DA-Data.xml", true)]
        [InlineData("DA004-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml", false)]
        [InlineData("DA005-SelectRowData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA006-SelectTestValue-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA007-SelectRepeatingData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA008-TableElementWithNoTable.docx", "DA-Data.xml", true)]
        [InlineData("DA009-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA010-InvalidXml.docx", "DA-Data.xml", true)]
        [InlineData("DA011-SchemaError.docx", "DA-Data.xml", true)]
        [InlineData("DA012-OtherMarkupTypes.docx", "DA-Data.xml", true)]
        [InlineData("DA013-Runs.docx", "DA-Data.xml", false)]
        [InlineData("DA014-TwoRuns-NoValuesSelected.docx", "DA-Data.xml", true)]
        [InlineData("DA015-TwoRunsXmlExceptionInFirst.docx", "DA-Data.xml", true)]
        [InlineData("DA016-TwoRunsSchemaErrorInSecond.docx", "DA-Data.xml", true)]
        [InlineData("DA017-FiveRuns.docx", "DA-Data.xml", true)]
        [InlineData("DA018-SmartQuotes.docx", "DA-Data.xml", false)]
        [InlineData("DA019-RunIsEntireParagraph.docx", "DA-Data.xml", false)]
        [InlineData("DA020-TwoRunsAndNoOtherContent.docx", "DA-Data.xml", true)]
        [InlineData("DA021-NestedRepeat.docx", "DA-DataNestedRepeat.xml", false)]
        [InlineData("DA022-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA023-RepeatWOEndRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA026-InvalidRootXmlElement.docx", "DA-Data.xml", true)]
        [InlineData("DA027-XPathErrorInPara.docx", "DA-Data.xml", true)]
        [InlineData("DA028-NoPrototypeRow.docx", "DA-Data.xml", true)]
        [InlineData("DA029-NoDataForCell.docx", "DA-Data.xml", true)]
        [InlineData("DA030-TooMuchDataForCell.docx", "DA-TooMuchDataForCell.xml", true)]
        [InlineData("DA031-CellDataInAttributes.docx", "DA-CellDataInAttributes.xml", true)]
        [InlineData("DA032-TooMuchDataForConditional.docx", "DA-TooMuchDataForConditional.xml", true)]
        [InlineData("DA033-ConditionalOnAttribute.docx", "DA-ConditionalOnAttribute.xml", false)]
        [InlineData("DA034-HeaderFooter.docx", "DA-Data.xml", false)]
        [InlineData("DA035-SchemaErrorInRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA036-SchemaErrorInConditional.docx", "DA-Data.xml", true)]

        [InlineData("DA100-TemplateDocument.docx", "DA-Data.xml", false)]
        [InlineData("DA101-TemplateDocument.docx", "DA-Data.xml", true)]
        [InlineData("DA102-TemplateDocument.docx", "DA-Data.xml", true)]

        [InlineData("DA201-TemplateDocument.docx", "DA-Data.xml", false)]
        [InlineData("DA202-TemplateDocument.docx", "DA-DataNotHighValueCust.xml", false)]
        [InlineData("DA203-Select-XPathFindsNoData.docx", "DA-Data.xml", true)]
        [InlineData("DA204-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml", false)]
        [InlineData("DA205-SelectRowData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA206-SelectTestValue-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA207-SelectRepeatingData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA209-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA210-InvalidXml.docx", "DA-Data.xml", true)]
        [InlineData("DA211-SchemaError.docx", "DA-Data.xml", true)]
        [InlineData("DA212-OtherMarkupTypes.docx", "DA-Data.xml", true)]
        [InlineData("DA213-Runs.docx", "DA-Data.xml", false)]
        [InlineData("DA214-TwoRuns-NoValuesSelected.docx", "DA-Data.xml", true)]
        [InlineData("DA215-TwoRunsXmlExceptionInFirst.docx", "DA-Data.xml", true)]
        [InlineData("DA216-TwoRunsSchemaErrorInSecond.docx", "DA-Data.xml", true)]
        [InlineData("DA217-FiveRuns.docx", "DA-Data.xml", true)]
        [InlineData("DA218-SmartQuotes.docx", "DA-Data.xml", false)]
        [InlineData("DA219-RunIsEntireParagraph.docx", "DA-Data.xml", false)]
        [InlineData("DA220-TwoRunsAndNoOtherContent.docx", "DA-Data.xml", true)]
        [InlineData("DA221-NestedRepeat.docx", "DA-DataNestedRepeat.xml", false)]
        [InlineData("DA222-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA223-RepeatWOEndRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA226-InvalidRootXmlElement.docx", "DA-Data.xml", true)]
        [InlineData("DA227-XPathErrorInPara.docx", "DA-Data.xml", true)]
        [InlineData("DA228-NoPrototypeRow.docx", "DA-Data.xml", true)]
        [InlineData("DA229-NoDataForCell.docx", "DA-Data.xml", true)]
        [InlineData("DA230-TooMuchDataForCell.docx", "DA-TooMuchDataForCell.xml", true)]
        [InlineData("DA231-CellDataInAttributes.docx", "DA-CellDataInAttributes.xml", true)]
        [InlineData("DA232-TooMuchDataForConditional.docx", "DA-TooMuchDataForConditional.xml", true)]
        [InlineData("DA233-ConditionalOnAttribute.docx", "DA-ConditionalOnAttribute.xml", false)]
        [InlineData("DA234-HeaderFooter.docx", "DA-Data.xml", false)]
        [InlineData("DA235-Crashes.docx", "DA-Content-List.xml", false)]
        [InlineData("DA236-Page-Num-in-Footer.docx", "DA-Content-List.xml", false)]
        [InlineData("DA237-SchemaErrorInRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA238-SchemaErrorInConditional.docx", "DA-Data.xml", true)]
        [InlineData("DA239-RunLevelCC-Repeat.docx", "DA-Data.xml", false)]

        [InlineData("DA250-ConditionalWithRichXPath.docx", "DA250-Address.xml", false)]
        [InlineData("DA251-EnhancedTables.docx", "DA-Data.xml", false)]
        [InlineData("DA252-Table-With-Sum.docx", "DA-Data.xml", false)]
        [InlineData("DA253-Table-With-Sum-Run-Level-CC.docx", "DA-Data.xml", false)]
        [InlineData("DA254-Table-With-XPath-Sum.docx", "DA-Data.xml", false)]
        [InlineData("DA255-Table-With-XPath-Sum-Run-Level-CC.docx", "DA-Data.xml", false)]
        [InlineData("DA256-NoInvalidDocOnErrorInRun.docx", "DA-Data.xml", true)]
        [InlineData("DA257-OptionalRepeat.docx", "DA-Data.xml", false)]
        [InlineData("DA258-ContentAcceptsCharsAsXPathResult.docx", "DA-Data.xml", false)]
        [InlineData("DA259-MultiLineContents.docx", "DA-Data.xml", false)]
        [InlineData("DA260-RunLevelRepeat.docx", "DA-Data.xml", false)]
        [InlineData("DA261-RunLevelConditional.docx", "DA-Data.xml", false)]
        [InlineData("DA262-ConditionalNotMatch.docx", "DA-Data.xml", false)]
        [InlineData("DA263-ConditionalNotMatch.docx", "DA-DataSmallCustomer.xml", false)]
        [InlineData("DA264-InvalidRunLevelRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA265-RunLevelRepeatWithWhiteSpaceBefore.docx", "DA-Data.xml", false)]
        [InlineData("DA266-RunLevelRepeat-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA268-Block-Conditional-In-Table-Cell.docx", "DA268-data.xml", false)]

        public void DA101(string name, string data, bool err)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            FileInfo dataFile = new FileInfo(Path.Combine(sourceDir.FullName, data));

            WmlDocument wmlTemplate = new WmlDocument(templateDocx.FullName);
            XElement xmldata = XElement.Load(dataFile.FullName);

            bool returnedTemplateError;
            WmlDocument afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            afterAssembling.SaveAs(assembledDocx.FullName);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(afterAssembling.DocumentByteArray, 0, afterAssembling.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator v = new OpenXmlValidator();
                    var valErrors = v.Validate(wDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));

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

            Assert.Equal(err, returnedTemplateError);
            AssertValidLastParagraph(afterAssembling); // experimental
        }

        [Theory]
        [InlineData("DA259-MultiLineContents.docx", "DA-Data.xml", false)]
        public void DA259(string name, string data, bool err)
        {
            DA101(name, data, err);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            WmlDocument afterAssembling = new WmlDocument(assembledDocx.FullName);
            int brCount = afterAssembling.MainDocumentPart
                            .Element(W.body)
                            .Elements(W.p).ElementAt(1)
                            .Elements(W.r)
                            .Elements(W.br).Count();
            Assert.Equal(4, brCount);
        }

        [Fact]
        public void DA240()
        {
            string name = "DA240-Whitespace.docx";
            DA101(name, "DA240-Whitespace.xml", false);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            WmlDocument afterAssembling = new WmlDocument(assembledDocx.FullName);

            // when elements are inserted that begin or end with white space, make sure white space is preserved
            string firstParaTextIncorrect = afterAssembling.MainDocumentPart.Element(W.body).Elements(W.p).First().Value;
            Assert.Equal("Content may or may not have spaces: he/she; he, she; he and she.", firstParaTextIncorrect);
            // warning: XElement.Value returns the string resulting from direct concatenation of all W.t elements. This is fast but ignores
            // proper handling of xml:space="preserve" attributes, which Word honors when rendering content. Below we also check
            // the result of UnicodeMapper.RunToString, which has been enhanced to take xml:space="preserve" into account.
            string firstParaTextCorrect = InnerText(afterAssembling.MainDocumentPart.Element(W.body).Elements(W.p).First());
            Assert.Equal("Content may or may not have spaces: he/she; he, she; he and she.", firstParaTextCorrect);
        }

        [Theory]
        [InlineData("DA024-TrackedRevisions.docx", "DA-Data.xml")]
        public void DA102_Throws(string name, string data)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            FileInfo dataFile = new FileInfo(Path.Combine(sourceDir.FullName, data));

            WmlDocument wmlTemplate = new WmlDocument(templateDocx.FullName);
            XElement xmldata = XElement.Load(dataFile.FullName);

            bool returnedTemplateError;
            WmlDocument afterAssembling;
            Assert.Throws<OpenXmlPowerToolsException>(() =>
            {
                afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out returnedTemplateError);
            });
        }

        [Fact]
        public void DATemplateMaior()
        {
            // this test case was causing incorrect behavior of OpenXmlRegex when replacing fields in paragraphs that contained
            // lastRenderedPageBreak XML elements. Recent fixes relating to UnicodeMapper and OpenXmlRegex addressed it.
            string name = "DA-TemplateMaior.docx";
            DA101(name, "DA-templateMaior.xml", false);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            var afterAssembling = new WmlDocument(assembledDocx.FullName);

            var descendants = afterAssembling.MainDocumentPart.Value;

            Assert.False(descendants.Contains(">"), "Found > on text");
        }

        [Fact]
        public void DAXmlError()
        {
            /* The assembly below would originally (prior to bug fixes) cause an exception to be thrown during assembly: 
                 System.ArgumentException : '', hexadecimal value 0x01, is an invalid character.
             */
            string name = "DA-xmlerror.docx";
            string data = "DA-xmlerror.xml";

            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var dataFile = new FileInfo(Path.Combine(sourceDir.FullName, data));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmlData = XElement.Load(dataFile.FullName);

            var afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out bool returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            afterAssembling.SaveAs(assembledDocx.FullName);
        }

        [Theory]
        [InlineData("DA025-TemplateDocument.docx", "DA-Data.xml", false)]
        public void DA103_UseXmlDocument(string name, string data, bool err)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            FileInfo dataFile = new FileInfo(Path.Combine(sourceDir.FullName, data));

            WmlDocument wmlTemplate = new WmlDocument(templateDocx.FullName);
            XmlDocument xmldata = new XmlDocument();
            xmldata.Load(dataFile.FullName);

            bool returnedTemplateError;
            WmlDocument afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            afterAssembling.SaveAs(assembledDocx.FullName);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(afterAssembling.DocumentByteArray, 0, afterAssembling.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator v = new OpenXmlValidator();
                    var valErrors = v.Validate(wDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));
                    Assert.Empty(valErrors);
                }
            }

            Assert.Equal(err, returnedTemplateError);
        }

        [Theory]
        [InlineData("DA267-xmlerror.docx", "DA267-xmlerror.xml", true)]
        public void DA267_XmlError(string name, string data, bool expectError)
        {
            // this test docx template is invalid -- the footer, which contains fields with markup,
            // is itself wrapped in an invisible content control that should not be there.
            // Since that invisible outer content control contains formatted text and other content controls,
            // instead of containing actual XML, an error is thrown when we try to parse that content as xml.
            // This XML error is embedded in a run where the error occurred.

            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo templateDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
            FileInfo dataFile = new FileInfo(Path.Combine(sourceDir.FullName, data));

            WmlDocument wmlTemplate = new WmlDocument(templateDocx.FullName);
            XmlDocument xmldata = new XmlDocument();
            xmldata.Load(dataFile.FullName);

            bool returnedTemplateError;
            WmlDocument afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out returnedTemplateError);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            afterAssembling.SaveAs(assembledDocx.FullName);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(afterAssembling.DocumentByteArray, 0, afterAssembling.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    OpenXmlValidator v = new OpenXmlValidator();
                    var valErrors = v.Validate(wDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));
                    Assert.Empty(valErrors);
                }
            }

            Assert.Equal(expectError, returnedTemplateError);
        }

        private void AssertValidLastParagraph(WmlDocument doc)
        {
            // ensure the last block-level item in the document is a paragraph
            var lastInBody = doc.MainDocumentPart.Element(W.body).LastNode as XElement;
            Assert.NotNull(lastInBody);
            var lastBlock = lastInBody.PreviousNode as XElement;
            Assert.Equal(W.p, lastBlock.Name);
            // ... and that that paragraph does not have a section break
            Assert.Null(lastBlock.Descendants(W.sectPr).FirstOrDefault());
            // (because if it DID, that would mean the last section has zero paragraphs, which seems iffy --
            // Word never does that on its own, and DocumentBuilder mishandles it!)
        }

        [Fact]
        public void DA280()
        {
            string name = "DA280-ConditionalSectionBreaks.docx";
            DA101(name, "DA280-data.xml", false);
            var outfile = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, name.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            WmlDocument afterAssembling = new WmlDocument(outfile.FullName);
            AssertValidLastParagraph(afterAssembling);
        }

        // tests for indirect insert (legacy -- prior to introduction of DocumentComposer)
        private void DA300(string nameParent, string nameChild, string data, bool err, bool errChild)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo parentTemplateDocx = new FileInfo(Path.Combine(sourceDir.FullName, nameParent));
            FileInfo childTemplateDocx = new FileInfo(Path.Combine(sourceDir.FullName, nameChild));
            FileInfo dataFile = new FileInfo(Path.Combine(sourceDir.FullName, data));

            XElement xmldata = XElement.Load(dataFile.FullName);
            WmlDocument wmlParentTemplate = new WmlDocument(parentTemplateDocx.FullName);
            WmlDocument wmlChildTemplate = new WmlDocument(childTemplateDocx.FullName);

            bool returnedChildError;
            WmlDocument assembledChild = DocumentAssembler.AssembleDocument(wmlChildTemplate, xmldata, out returnedChildError);
            Assert.Equal(errChild, returnedChildError);

            bool returnedTemplateError;
            WmlDocument assembledParent = DocumentAssembler.AssembleDocument(wmlParentTemplate, xmldata, out returnedTemplateError);
            Assert.Equal(err, returnedTemplateError);

            List<Source> sources = new List<Source>()
            {
                new Source(assembledParent),
                new Source(assembledChild, nameChild),
            };
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName,
                parentTemplateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx")
            ));
            DocumentBuilder.BuildDocument(sources, assembledDocx.FullName);

            using (WordprocessingDocument wDoc = WordprocessingDocument.Open(assembledDocx.FullName, true))
            {
                OpenXmlValidator v = new OpenXmlValidator();
                var valErrors = v.Validate(wDoc).Where(ve => !s_BuilderExpectedErrors.Contains(ve.Description));
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

        [Fact]
        public void DA301_DocBuilder_Inserts()
        {
            var parent = "DA301-Insert.docx";
            var child = "DA301-Inserted.docx";
            var data = "DA301-Insert.xml";
            DA300(parent, child, data, false, false);
            var assembledDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, parent.Replace(".docx", "-processed-by-DocumentAssembler.docx")));
            WmlDocument afterAssembling = new WmlDocument(assembledDocx.FullName);
            var mainPart = afterAssembling.MainDocumentPart;
            var paragraphs = mainPart.Element(W.body).Elements(W.p);
            string para1Text = paragraphs.First().Value;
            string para2Text = paragraphs.ElementAt(1).Value;
            Assert.Equal("This is a parent document, Cheryl.", para1Text);
            Assert.Equal("Cheryl's bulleted item!", para2Text);
        }

        private static string InnerText(XContainer e)
        {
            return e.Descendants(W.r)
                .Where(r => r.Parent.Name != W.del)
                .Select(UnicodeMapper.RunToString)
                .StringConcatenate();
        }

        //private static bool AllCellsHaveParagraphs(WordprocessingDocument wordDoc)
        //{
        //    foreach (var part in wordDoc.ContentParts())
        //    {
        //        CellsInPartHaveParagraphs(part);
        //    }
        //}

        //private static void CellsInPartHaveParagraphs(OpenXmlPart part)
        //{
        //    XDocument xDoc = part.GetXDocument();

        //    var xDocRoot = RemoveGoBackBookmarks();

        //    // content controls in cells can surround the W.tc element, so transform so that such content controls are within the cell content
        //    xDocRoot = (XElement)NormalizeContentControlsInCells(xDoc.Root);

        //}

        //private static bool NodeCellsHaveParagraphs(XNode node)
        //{
        //    XElement element = node as XElement;
        //    if (element != null)
        //    {
        //        if (element.Name == W.tc)
        //        {
        //            element.Elem
        //            var childMeta = element.Elements().Where(n => n.Name == PtOpenXml.Insert);
        //            var count = childMeta.Count();
        //            if (count > 0)
        //            {
        //                var pAt = element.Attributes();
        //                var pPr = element.Elements(W.pPr).FirstOrDefault();
        //                XElement p = null;
        //                List<XElement> list = new List<XElement>();
        //                foreach (var elem in element.Elements().Where(e => e.Name != W.pPr))
        //                {
        //                    if (elem.Name == PtOpenXml.Insert)
        //                    {
        //                        if (p != null)
        //                        {
        //                            list.Add(p);
        //                            p = null;
        //                        }
        //                        list.Add(elem);
        //                    }
        //                    else
        //                    { // non-insert content
        //                        if (p == null)
        //                        {
        //                            p = new XElement(W.p, pAt, pPr);
        //                        }
        //                        p.Add(new XElement(elem.Name, elem.Attributes(), elem.Nodes().Select(n => FixDocBuilderInserts(n, te))));
        //                    }
        //                }
        //                if (p != null)
        //                {
        //                    list.Add(p);
        //                    p = null;
        //                }
        //                return list;
        //            }
        //        }
        //        return new XElement(element.Name,
        //            element.Attributes(),
        //            element.Nodes().Select(n => FixDocBuilderInserts(n, te)));
        //    }
        //    return node;
        //}

        private static List<string> s_ExpectedErrors = new List<string>()
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
            "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:name' has invalid value 'useWord2013TrackBottomHyphenation'. The Enumeration constraint failed.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:allStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:alternateStyleNames' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:clearFormatting' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:customStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnNumbering' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnParagraphs' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnRuns' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnTables' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:headingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:latentStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:numberingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:stylesInUse' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:tableStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:top3HeadingStyles' attribute is not declared.",
            "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:visibleStyles' attribute is not declared.",
            "The element has invalid child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:doNotEmbedSmartTags'.",
        };

        private static List<string> s_BuilderExpectedErrors = new List<string>()
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
        };
    }
}

#endif
