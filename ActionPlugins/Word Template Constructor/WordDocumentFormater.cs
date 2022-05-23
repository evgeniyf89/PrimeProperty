using DocumentFormat.OpenXml.Wordprocessing;
using SoftLine.ActionPlugins.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using SoftLine.ActionPlugins.SharePoint;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Xml.XPath;
using System.Xml;

namespace SoftLine.ActionPlugins.WordTemplateConstructor
{
    public class WordDocumentFormater : IWordDocumentFormater
    {

        public byte[] Form(byte[] wordTemplate, PrintFormSettings settings)
        {
            using (var templateStream = new MemoryStream())
            {
                templateStream.Write(wordTemplate, 0, wordTemplate.Length);
                using (var word = WordprocessingDocument.Open(templateStream, true))
                {
                    ProcessDocumentVariables(word, settings);
                }
                return templateStream.ToArray();
            }
        }

        private void GenerateDrawing(string relationshipId, BookmarkStart bookmarkStart)
        {
            var drawing1 = new Drawing();
            DW.Anchor anchor1 = new DW.Anchor()
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)114300U,
                DistanceFromRight = (UInt32Value)114300U,
                SimplePos = false,
                RelativeHeight = (UInt32Value)251658240U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };
            DW.SimplePosition simplePosition1 = new DW.SimplePosition() { X = 0L, Y = 0L };

            DW.HorizontalPosition horizontalPosition1 = new DW.HorizontalPosition()
            {
                RelativeFrom = DW.HorizontalRelativePositionValues.Character
            };
            DW.PositionOffset positionOffset1 = new DW.PositionOffset
            {
                Text = "2515"
            };

            horizontalPosition1.Append(positionOffset1);

            DW.VerticalPosition verticalPosition1 = new DW.VerticalPosition() { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph };
            DW.PositionOffset positionOffset2 = new DW.PositionOffset
            {
                Text = "-3200"
            };

            verticalPosition1.Append(positionOffset2);
            DW.Extent extent1 = new DW.Extent() { Cx = (Int64Value)(2857500L / 1.3), Cy = (Int64Value)(1609725L / 1.3) };
            DW.EffectExtent effectExtent1 = new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 9525L };
            DW.WrapNone wrapNone1 = new DW.WrapNone();
            DW.DocProperties docProperties1 = new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Рисунок 1" };

            DW.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new DW.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            PIC.Picture picture1 = new PIC.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            PIC.NonVisualPictureProperties nonVisualPictureProperties1 = new PIC.NonVisualPictureProperties();
            PIC.NonVisualDrawingProperties nonVisualDrawingProperties1 = new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "фото 1.jpg" };
            PIC.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new PIC.NonVisualPictureDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            PIC.BlipFill blipFill1 = new PIC.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = relationshipId };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            PIC.ShapeProperties shapeProperties1 = new PIC.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = (Int64Value)(2857500L / 1.3), Cy = (Int64Value)(1609725L / 1.3) };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Page };
            Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
            percentageWidth1.Text = "0";

            relativeWidth1.Append(percentageWidth1);

            Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
            Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
            percentageHeight1.Text = "0";

            relativeHeight1.Append(percentageHeight1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);

            drawing1.Append(anchor1);
            bookmarkStart.Parent.InsertAfter(new DocumentFormat.OpenXml.Math.Run(drawing1), bookmarkStart);
        }

        private void ProcessDocumentVariables(WordprocessingDocument wordprocessingDocument, PrintFormSettings formSettings)
        {
            var settings = wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
            var DOCVARIABLE_POSTFIX = @"_{0}";
            var updateFields = new UpdateFieldsOnOpen
            {
                Val = new OnOffValue(true)
            };
            settings.PrependChild(updateFields);

            if (settings.Descendants<DocumentVariables>().Count() == 0)
            {
                settings.Append(new DocumentVariables());
            }

            var documentVariables = settings.Descendants<DocumentVariables>().First();

            #region Check docvariables in template and provided values
            var fieldsToCheck = new List<string>();
            var fields = GetFieldDefinitions(formSettings, wordprocessingDocument.MainDocumentPart.Document, true);
            foreach (var field in fields)
            {
                var fieldName = (GetDocumentVariableName(field.Key) ?? String.Empty).Trim('\"');
                if (!CanProcessDocumentVariable(formSettings, fieldName))
                {
                    fieldsToCheck.Add(fieldName);
                }
            }
            #endregion

            #region Process document variables from values dictionary. Process document variables with format info.


            foreach (var parameter in formSettings.Parameters)
            {
                if (parameter.Value is List<Dictionary<string, object>>)
                {
                    var tableRows = parameter.Value as List<Dictionary<string, object>>;

                    for (int index = 0; index < tableRows.Count; index++)
                    {
                        foreach (string columnName in tableRows[index].Keys)
                        {
                            AddDocumentVariable(
                                documentVariables,
                                columnName + string.Format(DOCVARIABLE_POSTFIX, index),
                                tableRows[index][columnName]);
                        }
                    }
                }
                else
                {
                    if (parameter.FieldType == OptionSets.DocumentParameterFieldType.Image)
                    {
                        var imagePart = wordprocessingDocument.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
                        var image = parameter.Value as byte[];
                        using (var stream = new MemoryStream(image))
                        {
                            imagePart.FeedData(stream);
                        }
                        var bookmarkStarts = wordprocessingDocument.MainDocumentPart.RootElement.Descendants<BookmarkStart>().ToArray();
                        var relationshipId = wordprocessingDocument.MainDocumentPart.GetIdOfPart(imagePart);
                        foreach (var bookmarkStart in bookmarkStarts.Where(x => x.Name == parameter.WordDocVariable))
                        {
                            GenerateDrawing(relationshipId, bookmarkStart);
                            bookmarkStart.Remove();
                        }
                    }
                    else
                        AddDocumentVariable(documentVariables, parameter.WordDocVariable, parameter.Value);
                }
            }


            foreach (KeyValuePair<string, List<FieldCode>> keyValuePair in
                GetFieldDefinitions(formSettings, wordprocessingDocument.MainDocumentPart.Document.Body).Where(fd => fd.Key.Contains(':')))
            {
                string documentVariableName = GetDocumentVariableName(keyValuePair.Key);
                string documentVariableNameWithoutFormatInfo = documentVariableName.Substring(0, documentVariableName.IndexOf(':'));
                var variableValue = formSettings.Parameters
                    .FirstOrDefault(x => x.WordDocVariable == documentVariableNameWithoutFormatInfo)?.Value;
                AddDocumentVariable(documentVariables, documentVariableName, variableValue);
            }

            #endregion
            foreach (var table in wordprocessingDocument.MainDocumentPart.Document.Descendants<Table>())
            {

                var tableIdOpenXmlElement = table.PreviousSibling();

                if (tableIdOpenXmlElement == null)
                    continue;

                var tableIdFieldDefinition = GetFieldDefinitions(formSettings, tableIdOpenXmlElement);

                string tableId = null;

                if (tableIdFieldDefinition.Count == 0)
                {
                    continue;
                }
                else if (tableIdFieldDefinition.Count == 1)
                {
                    tableId = GetDocumentVariableName(tableIdFieldDefinition.Keys.First());

                    if (!CanProcessDocumentVariable(formSettings, tableId))
                        continue;
                }
                tableIdOpenXmlElement.Remove();
                var tableValue = formSettings.Parameters
                    .FirstOrDefault(x => x.WordDocVariable == tableId)?.Value as List<Dictionary<string, object>>;
                foreach (var tableRow in table.Descendants<TableRow>())
                {
                    var fieldsDefinition = GetFieldDefinitions(formSettings, tableRow);

                    if (fieldsDefinition.Count == 0)
                    {
                        continue;
                    }
                    var canProcessRow = true;


                    if (tableValue == null)
                        continue;

                    foreach (var fieldDefinition in fieldsDefinition.Keys)
                    {
                        var fieldName = GetDocumentVariableName(fieldDefinition);

                        if (tableValue.FirstOrDefault(tr => tr.Keys.Contains(fieldName)) == null)
                        {
                            canProcessRow = false;
                            break;
                        }
                    }

                    if (!canProcessRow)
                        continue;

                    var newRowsCount = tableValue.Count;

                    var lastRow = tableRow;

                    for (int newRowIndex = 1; newRowIndex < newRowsCount; newRowIndex++)
                    {
                        var newTableRow = tableRow.Clone() as TableRow;

                        // update all fields in new row by adding postfix "_[newRowIndex]" for all document variables
                        AppendDocumentVariableIndices(formSettings, newTableRow, tableId, newRowIndex);
                        lastRow.InsertAfterSelf(newTableRow);
                        lastRow = newTableRow;
                    }

                    // update all fields of first row (newIndex = 0)
                    fieldsDefinition = GetFieldDefinitions(formSettings, tableRow);

                    AppendDocumentVariableIndices(formSettings, tableRow, tableId, 0);

                    break;
                }
            }



        }

        private Dictionary<string, List<FieldCode>> GetFieldDefinitions(PrintFormSettings formSettings, OpenXmlElement container, bool dontCheckDocumentVariable = false)
        {
            Dictionary<string, List<FieldCode>> fieldDefinitions = new Dictionary<string, List<FieldCode>>();
            List<FieldCode> fieldDefinitionParts = new List<FieldCode>();
            StringBuilder fieldDefinitionStringBuilder = new StringBuilder();
            bool isConcatenatingFieldDefinition = false;

            // view all FieldChar and FieldCode elements for concatenate field definitions and determine document variable names inside them
            foreach (OpenXmlElement fieldDefinitionPart in container.Descendants().Where(fc =>
                fc is FieldChar || fc is FieldCode))
            {
                if (fieldDefinitionPart is FieldChar)
                {
                    switch ((fieldDefinitionPart as FieldChar).FieldCharType.Value)
                    {
                        case FieldCharValues.Begin:
                            // start of field definition
                            isConcatenatingFieldDefinition = true;
                            break;

                        case FieldCharValues.Separate:
                        case FieldCharValues.End:
                            // end of field definition
                            isConcatenatingFieldDefinition = false;
                            break;
                    }
                }

                if (isConcatenatingFieldDefinition && fieldDefinitionPart is FieldCode)
                {
                    fieldDefinitionStringBuilder.Append((fieldDefinitionPart as FieldCode).InnerText);
                    fieldDefinitionParts.Add(fieldDefinitionPart as FieldCode);
                }

                if (!isConcatenatingFieldDefinition && fieldDefinitionParts.Count != 0)
                {
                    string fieldDefinitionFullText = fieldDefinitionStringBuilder.ToString();
                    string variableName = GetDocumentVariableName(fieldDefinitionFullText);

                    // process only document variables that exist in data table
                    if (variableName != null && (CanProcessDocumentVariable(formSettings, variableName) || dontCheckDocumentVariable))
                    {
                        if (fieldDefinitions.ContainsKey(fieldDefinitionFullText))
                            fieldDefinitions[fieldDefinitionFullText] = fieldDefinitionParts;
                        else
                            fieldDefinitions.Add(fieldDefinitionFullText, fieldDefinitionParts);
                    }

                    fieldDefinitionStringBuilder = new StringBuilder();
                    fieldDefinitionParts = new List<FieldCode>();

                    continue;
                }
            }

            return fieldDefinitions;
        }

        private bool CanProcessDocumentVariable(PrintFormSettings formSettings, string documentVariableName)
        {
            if (documentVariableName.Contains(':'))
            {
                documentVariableName = documentVariableName.Substring(0, documentVariableName.IndexOf(':'));
            }
            if (formSettings.Parameters.Select(x => x.WordDocVariable).Contains(documentVariableName))
            {
                return true;
            }
            foreach (var tableVariable in formSettings.Parameters.Select(x => x.Value).OfType<List<Dictionary<string, object>>>())
            {
                var table = tableVariable.FirstOrDefault();
                if (table == null)
                    continue;
                if (table.Keys.Contains(documentVariableName))
                    return true;
            }
            return false;
        }

        private void AppendDocumentVariableIndices(PrintFormSettings formSettings, TableRow tableRow, string tableId, int newIndex)
        {
            string DOCVARIABLE_POSTFIX = @"_{0}";
            Dictionary<string, List<FieldCode>> fieldDefinitions = GetFieldDefinitions(formSettings, tableRow);

            foreach (string fieldDefinitionFullText in fieldDefinitions.Keys)
            {
                // FieldCode contains end of the document variable name
                FieldCode documentVariableLastFieldCode = null;
                string documentVariableName = GetDocumentVariableName(fieldDefinitionFullText);

                // process only document variables that exist in data table
                if (documentVariableName != null && CanProcessDocumentVariable(formSettings, documentVariableName))
                {
                    // index of last symbol of document variable in field definition
                    int documentVariableEndPosition = fieldDefinitionFullText.IndexOf(documentVariableName) + documentVariableName.Length;

                    // determine last FieldCode contains end of the name of document variable
                    int innerTextLength = 0;

                    foreach (FieldCode fieldCode in fieldDefinitions[fieldDefinitionFullText])
                    {
                        innerTextLength += fieldCode.InnerText.Length;

                        if (innerTextLength >= documentVariableEndPosition)
                        {
                            documentVariableLastFieldCode = fieldCode;

                            break;
                        }
                    }

                    if (documentVariableLastFieldCode != null)
                    {
                        // determine last symbol of the document variable name
                        int slashPosition = documentVariableLastFieldCode.Text.IndexOf(@"\");

                        if (slashPosition == -1)
                        {
                            slashPosition = documentVariableLastFieldCode.Text.Length;
                        }

                        int endVariableNamePosition =
                            documentVariableLastFieldCode.Text.Substring(0, slashPosition).Reverse().
                            TakeWhile(c => !char.IsLetterOrDigit(c)).Count();

                        endVariableNamePosition =
                            documentVariableLastFieldCode.Text.Substring(0, slashPosition).Length -
                            endVariableNamePosition;

                        // append new index to a document variable
                        documentVariableLastFieldCode.Text =
                            documentVariableLastFieldCode.Text.Insert(endVariableNamePosition, string.Format(DOCVARIABLE_POSTFIX, newIndex));
                    }
                }
            }
        }

        private string GetDocumentVariableName(string fieldDefinitionFullText)
        {
            var DocVariable = @"DOCVARIABLE";
            if (!fieldDefinitionFullText.Contains(DocVariable))
            {
                return null;
            }

            var startIndex = fieldDefinitionFullText.IndexOf(DocVariable) + DocVariable.Length;
            var len = fieldDefinitionFullText.IndexOf(@"\") - DocVariable.Length - fieldDefinitionFullText.IndexOf(DocVariable);
            if (startIndex < 0 || len < 0)
                return null;

            return fieldDefinitionFullText.Substring(startIndex, len).Trim();
        }

        private void AddDocumentVariable(DocumentVariables documentVariables, string variableName, object variableValue)
        {
            var documentVariable = documentVariables.OfType<DocumentVariable>().FirstOrDefault(dv => dv.Name == variableName);

            if (documentVariable == null)
            {
                documentVariable = new DocumentVariable { Name = variableName };

                documentVariables.Append(documentVariable);
            }
            if (variableName == "ОсновныеПреимущества")
            {
                var byteArray = Encoding.ASCII.GetBytes(variableValue as string);
                var stream = new MemoryStream(byteArray);

                var htmlDoc = new XPathDocument(stream);

                var navigator = htmlDoc.CreateNavigator();
                var mngr = new XmlNamespaceManager(navigator.NameTable);
                mngr.AddNamespace("xhtml", "http://www.w3.org/1999/xhtml");

                var ni = navigator.Select("//xhtml:p", mngr);
                while (ni.MoveNext())
                {
                    documentVariable.AppendChild(new Paragraph(new Run(new Text(ni.Current.Value))));
                }
            }
            else
                documentVariable.Val = GetFormattedDocumentVariableValue(variableName, variableValue);
        }

        private string GetFormattedDocumentVariableValue(string variableName, object sourceValue)
        {
            string NULL_VARIABLE_VALUE = null;
            string FORMAT_DATEFULL = null;
            string FORMAT_SUMINWORDS = null;
            if (sourceValue == null)
                return NULL_VARIABLE_VALUE;

            if (sourceValue is DateTime)
            {
                return variableName.ToLower().Contains(FORMAT_DATEFULL ?? string.Empty) ?
                    ((DateTime)sourceValue).ToLongDateString() :
                    ((DateTime)sourceValue).ToShortDateString();
            }

            if (sourceValue is double)
            {
                return ((double)sourceValue).ToString("n");
            }

            if (sourceValue is int)
                return ((int)sourceValue).ToString("n0");

            var stringValue = sourceValue.ToString();

            return !string.IsNullOrEmpty(stringValue) ? stringValue : NULL_VARIABLE_VALUE;
        }

        public void ConvertHTML(string htmlFileName, string docFileName)
        {
            // Create a Wordprocessing document. 
            using (WordprocessingDocument package = WordprocessingDocument.Create(docFileName, WordprocessingDocumentType.Document))
            {
                // Add a new main document part. 
                package.AddMainDocumentPart();

                // Create the Document DOM. 
                package.MainDocumentPart.Document = new Document(new Body());
                Body body = package.MainDocumentPart.Document.Body;
                var htmlText = "<ul class=\"ul-styled\"><li>landmark architecture from one of London’s leading architects;</li><li>500 meters from sandy beach and city amenities of Limassol’s tourist area;</li><li>sea views from floors 6 and 7;</li><li>extended hotel-type facilities: outdoor pool, half-Olympic size indoor heated pool, SPA, concierge, playroom, club house, underground parking, large outdoor gardens, gym, playground for kids;</li><li>penthouses with private pools;</li><li>apartments with private gardens;</li><li>spacious duplexes and triplexes serving as a good alternative to a villa;</li><li>high ceilings (3.15 m);</li><li>high standard finishes (parquet floors, high doors of 2.4 m, security entrance doors, thermal aluminum window frames, top standard in-built furniture, and sanitary ware);</li><li>water underfloor heating andVRV air conditioning.</li></ul>";
                var byteArray = Encoding.ASCII.GetBytes(htmlText);
                var stream = new MemoryStream(byteArray);

                var htmlDoc = new XPathDocument(stream);

                var navigator = htmlDoc.CreateNavigator();
                var mngr = new XmlNamespaceManager(navigator.NameTable);
                mngr.AddNamespace("xhtml", "http://www.w3.org/1999/xhtml");

                var ni = navigator.Select("//xhtml:p", mngr);
                while (ni.MoveNext())
                {
                    body.AppendChild(new Paragraph(new Run(new Text(ni.Current.Value))));
                }
                package.MainDocumentPart.Document.Save();
            }
        }

    }
}
