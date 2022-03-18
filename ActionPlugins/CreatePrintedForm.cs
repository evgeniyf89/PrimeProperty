using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using HtmlAgilityPack;

namespace SoftLine.ActionPlugins
{
    public class CreatePrintedForm
    {
        public void RetrivePrintedForm()
        {
            var path = @"E:\Аренда прайслист разметка.docx";
            var imageFilename = @"C:\Users\Fedotoveni\Desktop\фото 1.jpg";
            byte[] byteArray = File.ReadAllBytes(path);
            var nameValues = new Dictionary<string, object>()
            {
                ["Номер"] = "702-1A",
                ["Пк"] = 10479,
                ["Этаж"] = 7,
                ["КонецСтроительства"] = DateTime.Now,
                ["КолвоСпален"] = 2,
                ["КолвоПарковок"] = 1,
                ["ВнутренПлощадь"] = 124.7,
                ["КрВерандаПлощадь"] = 31.3,
                ["ОткрВерандаПлощадь"] = 134.1,
                ["УчастокПлощадь"] = 0.0,
                ["ОбщаяПолзовПлощадь"] = 6.4,
                ["ОбщаяПлощадь"] = 296.4,
                ["РастояниеДоМоря"] = 200
            };
            using (var templateStream = new MemoryStream())
            {
                templateStream.Write(byteArray, 0, byteArray.Length);
                using (var word = WordprocessingDocument.Open(templateStream, true))
                {
                    ProcessDocumentVariables(word, nameValues);
                    var imagePart = word.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
                    using (FileStream stream = new FileStream(imageFilename, FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }
                    var bookmarkStarts = word.MainDocumentPart.RootElement.Descendants<BookmarkStart>().ToArray();
                    var relationshipId = word.MainDocumentPart.GetIdOfPart(imagePart);
                    foreach (var bookmarkStart in bookmarkStarts.Where(x => x.Name == "Image"))
                    {
                        GenerateDrawing(relationshipId, bookmarkStart);
                        bookmarkStart.Remove();
                    }
                }
                var savePath = @"E:\Аренда прайслист разметка c параметрами.docx";
                File.WriteAllBytes(savePath, templateStream.ToArray());
            }

        }

        public void RenderHtmlToPdf()
        {
            var imagePath = @"C:\Users\Fedotoveni\Desktop\фото 1.jpg";
            var imageArray = File.ReadAllBytes(imagePath);
            var base64ImageRepresentation = "data:image/jpeg;base64," + Convert.ToBase64String(imageArray);
            var path = @"C:/Users/Fedotoveni/Downloads/SodaPDF-converted-Arenda_prayslist/OutDocument/pg_0001.htm";
            //var html = File.ReadAllText(path);
            var doc = new HtmlDocument();
            doc.Load(path);
            var imageid = doc.GetElementbyId("mainPicture");
            var src = imageid.Attributes["src"];
            if (src is null)
                imageid.Attributes.Add("src", base64ImageRepresentation);
            else
                src.Value = base64ImageRepresentation;
        
            var newDoc = doc.DocumentNode.WriteTo();
            doc.Save(path);
        }

        public static void InsertImageIntoBookmark(WordprocessingDocument doc, BookmarkStart bookmarkStart, ImagePart imagePart)
        {
            // Remove anything present inside the bookmark
            OpenXmlElement elem = bookmarkStart.NextSibling();
            while (elem != null && !(elem is BookmarkEnd))
            {
                OpenXmlElement nextElem = elem.NextSibling();
                elem.Remove();
                elem = nextElem;
            }
            AddImageToBody(doc.MainDocumentPart.GetIdOfPart(imagePart), bookmarkStart);
        }

        public static ImagePart AddImagePart(MainDocumentPart mainPart, string imageFilename)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(imageFilename, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            return imagePart;
        }

        private static void AddImageToBody(string relationshipId, BookmarkStart bookmarkStart)
        {
            // Define the reference of the image.
            var anchor1 = new DW.Anchor()
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)114300U,
                DistanceFromRight = (UInt32Value)114300U,
                SimplePos = false,
                RelativeHeight = (UInt32Value)251658240U,
                BehindDoc = true,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };
            var element =
             new Drawing(
                 new DW.Inline(
                     new DW.Extent() { Cx = 990000L, Cy = 792000L },
                     new DW.EffectExtent()
                     {
                         LeftEdge = 0L,
                         TopEdge = 0L,
                         RightEdge = 0L,
                         BottomEdge = 0L
                     },
                     new DW.DocProperties()
                     {
                         Id = (UInt32Value)1U,
                         Name = "Picture 1"
                     },
                     new DW.NonVisualGraphicFrameDrawingProperties(
                         new A.GraphicFrameLocks() { NoChangeAspect = true }),
                     new A.Graphic(
                         new A.GraphicData(
                             new PIC.Picture(
                                 new PIC.NonVisualPictureProperties(
                                     new PIC.NonVisualDrawingProperties()
                                     {
                                         Id = (UInt32Value)0U,
                                         Name = "New Bitmap Image.jpg"
                                     },
                                     new PIC.NonVisualPictureDrawingProperties()),
                                 new PIC.BlipFill(
                                     new A.Blip(
                                         new A.BlipExtensionList(
                                             new A.BlipExtension()
                                             {
                                                 Uri =
                                                    "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                             })
                                     )
                                     {
                                         Embed = relationshipId,
                                         CompressionState =
                                         A.BlipCompressionValues.Print
                                     },
                                     new A.Stretch(
                                         new A.FillRectangle())),
                                 new PIC.ShapeProperties(
                                     new A.Transform2D(
                                         new A.Offset() { X = 0L, Y = 0L },
                                         new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                     new A.PresetGeometry(
                                         new A.AdjustValueList()
                                     )
                                     { Preset = A.ShapeTypeValues.Rectangle }))
                         )
                         { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                 )
                 {
                     DistanceFromTop = (UInt32Value)0U,
                     DistanceFromBottom = (UInt32Value)0U,
                     DistanceFromLeft = (UInt32Value)0U,
                     DistanceFromRight = (UInt32Value)0U,
                     EditId = "50D07946"
                 });
            element.Append(anchor1);
            bookmarkStart.Parent.InsertAfter(new Run(element), bookmarkStart);
        }

        public void GenerateDrawing(string relationshipId, BookmarkStart bookmarkStart)
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
            bookmarkStart.Parent.InsertAfter(new Run(drawing1), bookmarkStart);
        }

        public DocumentVariable GetVariableByName(string name, WordprocessingDocument document)
        {
            var documentSettings = document
                .MainDocumentPart
                .DocumentSettingsPart
                .Settings;
            var updateFields = new UpdateFieldsOnOpen
            {
                Val = new OnOffValue(true)
            };

            documentSettings.PrependChild(updateFields);

            if (documentSettings.Descendants<DocumentVariables>().Count() == 0)
            {
                documentSettings.Append(new DocumentVariables());
            }
            DocumentVariables documentVariables = documentSettings.Descendants<DocumentVariables>().First();
            var el = documentSettings
                .Elements<DocumentVariables>()
                .FirstOrDefault();
            return el?.Elements<DocumentVariable>()
                .FirstOrDefault(v => v.Name == name);
        }

        public ILookup<string, DocumentVariable> GetAllDocumentVariables(WordprocessingDocument document)
        {
            var documentSettings = document
                .MainDocumentPart
                .DocumentSettingsPart
                .Settings;
            var updateFields = new UpdateFieldsOnOpen
            {
                Val = new OnOffValue(true)
            };


            documentSettings.PrependChild(updateFields);

            if (documentSettings.Descendants<DocumentVariables>().Count() == 0)
            {
                documentSettings.Append(new DocumentVariables());
            }

            return documentSettings
                .Descendants<DocumentVariables>()
                .First()
                .Elements<DocumentVariable>()
                .ToLookup(x => x.Name.Value, y => y);
        }

        private void ProcessDocumentVariables(WordprocessingDocument wordprocessingDocument, Dictionary<string, object> valuesDictionary)
        {
            Settings settings = wordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings;
            string DOCVARIABLE_POSTFIX = @"_{0}";
            // Create object to update fields on open
            UpdateFieldsOnOpen updateFields = new UpdateFieldsOnOpen();
            updateFields.Val = new DocumentFormat.OpenXml.OnOffValue(true);

            // Insert object into settings part.
            settings.PrependChild<UpdateFieldsOnOpen>(updateFields);

            if (settings.Descendants<DocumentVariables>().Count() == 0)
            {
                settings.Append(new DocumentVariables());
            }

            DocumentVariables documentVariables = settings.Descendants<DocumentVariables>().First();

            #region Check docvariables in template and provided values
            var fieldsToCheck = new List<string>();
            var fields = GetFieldDefinitions(valuesDictionary, wordprocessingDocument.MainDocumentPart.Document, true);
            foreach (var field in fields)
            {
                var fieldName = (GetDocumentVariableName(field.Key) ?? String.Empty).Trim('\"');
                if (!CanProcessDocumentVariable(valuesDictionary, fieldName))
                {
                    fieldsToCheck.Add(fieldName);
                }
            }
            if (fieldsToCheck.Count > 0)
                throw new Exception("Для некоторых полей шаблона не найдены значения в CRM. Проверьте следующие параметры документов: " + String.Join(", ", fieldsToCheck));
            #endregion

            #region Process document variables from values dictionary. Process document variables with format info.

            // add document variables from values dictionary
            foreach (string variableName in valuesDictionary.Keys)
            {
                if (valuesDictionary[variableName] is List<Dictionary<string, object>>)
                {
                    var tableRows = valuesDictionary[variableName] as List<Dictionary<string, object>>;

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
                    AddDocumentVariable(documentVariables, variableName, valuesDictionary[variableName]);
                }
            }

            // add formatted document variables from document source code
            foreach (KeyValuePair<string, List<FieldCode>> keyValuePair in
                GetFieldDefinitions(valuesDictionary, wordprocessingDocument.MainDocumentPart.Document.Body).Where(fd => fd.Key.Contains(':')))
            {
                string documentVariableName = GetDocumentVariableName(keyValuePair.Key);
                string documentVariableNameWithoutFormatInfo = documentVariableName.Substring(0, documentVariableName.IndexOf(':'));

                AddDocumentVariable(documentVariables, documentVariableName, valuesDictionary[documentVariableNameWithoutFormatInfo]);
            }

            #endregion

            #region Process document variables stored in headers and footers
            /*foreach (var header in wordprocessingDocument.MainDocumentPart.HeaderParts)
            {
                foreach (var fc in header.RootElement.Descendants<FieldCode>())
                {
                    var fieldName = GetDocumentVariableName(fc.Text).Trim('\"');
                    if (valuesDictionary.ContainsKey(fieldName))
                    {
                        var value = GetFormattedDocumentVariableValue(fieldName, valuesDictionary[fieldName]);
                        header.RootElement.ReplaceChild<FieldCode>(new Text(value), fc);
                    }
                    
                    if (p.InnerXml.Contains(WordConst.DOCVARIABLE))
                    {
                        foreach (var value in valuesDictionary)
                        {
                            var fieldDefinition = String.Format("DOCVARIABLE  \"{0}\"  \\* MERGEFORMAT", value.Key);
                            if (p.InnerXml.Contains(fieldDefinition))
                            {
                                p.InnerXml = p.InnerXml.Replace(fieldDefinition, GetFormattedDocumentVariableValue(value.Key, value.Value));
                            }
                        }
                    }
                }
            }*/
            #endregion

            #region Process document variables stored in tables: add new rows and append indicies for document variables

            foreach (Table table in wordprocessingDocument.MainDocumentPart.Document.Descendants<Table>())
            {
                // get tableId container
                OpenXmlElement tableIdOpenXmlElement = table.PreviousSibling();

                if (tableIdOpenXmlElement == null)
                    continue;

                Dictionary<string, List<FieldCode>> tableIdFieldDefinition = GetFieldDefinitions(valuesDictionary, tableIdOpenXmlElement);

                string tableId = null;

                if (tableIdFieldDefinition.Count == 0)
                {
                    continue;
                }
                else if (tableIdFieldDefinition.Count == 1)
                {
                    tableId = GetDocumentVariableName(tableIdFieldDefinition.Keys.First());

                    if (!CanProcessDocumentVariable(valuesDictionary, tableId))
                        continue;
                }

                // remove tableId container
                tableIdOpenXmlElement.Remove();

                foreach (TableRow tableRow in table.Descendants<TableRow>())
                {
                    // don't process table rows without processable document variables
                    Dictionary<string, List<FieldCode>> fieldsDefinition = GetFieldDefinitions(valuesDictionary, tableRow);

                    if (fieldsDefinition.Count == 0)
                    {
                        continue;
                    }

                    // don't process table row if all its fields are not columns of table variable
                    var canProcessRow = true;
                    var tableValue = valuesDictionary[tableId] as List<Dictionary<string, object>>;

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

                    // determine count of new rows by count of list<object> corresponding to document variables in table row
                    int newRowsCount = (valuesDictionary[tableId] as List<Dictionary<string, object>>).Count();

                    // add new rows
                    TableRow lastRow = tableRow;

                    for (int newRowIndex = 1; newRowIndex < newRowsCount; newRowIndex++)
                    {
                        TableRow newTableRow = tableRow.Clone() as TableRow;

                        // update all fields in new row by adding postfix "_[newRowIndex]" for all document variables
                        AppendDocumentVariableIndices(valuesDictionary, newTableRow, tableId, newRowIndex);
                        lastRow.InsertAfterSelf(newTableRow);
                        lastRow = newTableRow;
                    }

                    // update all fields of first row (newIndex = 0)
                    fieldsDefinition = GetFieldDefinitions(valuesDictionary, tableRow);

                    AppendDocumentVariableIndices(valuesDictionary, tableRow, tableId, 0);

                    break;
                }
            }

            #endregion

        }

        private Dictionary<string, List<FieldCode>> GetFieldDefinitions(Dictionary<string, object> valuesDictionary, OpenXmlElement container, bool dontCheckDocumentVariable = false)
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
                    if (variableName != null && (CanProcessDocumentVariable(valuesDictionary, variableName) || dontCheckDocumentVariable))
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

        private bool CanProcessDocumentVariable(Dictionary<string, object> valuesDictionary, string documentVariableName)
        {
            if (documentVariableName.Contains(':'))
            {
                documentVariableName = documentVariableName.Substring(0, documentVariableName.IndexOf(':'));
            }
            if (valuesDictionary.Keys.Contains(documentVariableName))
            {
                return true;
            }
            foreach (var tableVariable in valuesDictionary.Values.OfType<List<Dictionary<string, object>>>())
            {
                var table = tableVariable.FirstOrDefault();
                if (table == null)
                    continue;
                if (table.Keys.Contains(documentVariableName))
                    return true;
            }
            return false;
        }

        private void AppendDocumentVariableIndices(Dictionary<string, object> valuesDictionary, TableRow tableRow, string tableId, int newIndex)
        {
            string DOCVARIABLE_POSTFIX = @"_{0}";
            Dictionary<string, List<FieldCode>> fieldDefinitions = GetFieldDefinitions(valuesDictionary, tableRow);

            foreach (string fieldDefinitionFullText in fieldDefinitions.Keys)
            {
                // FieldCode contains end of the document variable name
                FieldCode documentVariableLastFieldCode = null;
                string documentVariableName = GetDocumentVariableName(fieldDefinitionFullText);

                // process only document variables that exist in data table
                if (documentVariableName != null && CanProcessDocumentVariable(valuesDictionary, documentVariableName))
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

        private static string GetDocumentVariableName(string fieldDefinitionFullText)
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
            DocumentVariable documentVariable = documentVariables.OfType<DocumentVariable>().FirstOrDefault(dv => dv.Name == variableName);

            if (documentVariable == null)
            {
                documentVariable = new DocumentVariable { Name = variableName };

                documentVariables.Append(documentVariable);
            }

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

        private string FormatText(double sum, bool isMoney)
        {
            var res = string.Empty;

            var billions = new[]
                               {
                                   "миллиардов", "миллиард", "миллиарда", "миллиарда", "миллиарда", "миллиардов", "миллиардов",
                                   "миллиардов", "миллиардов", "миллиардов"
                               };
            var millions = new[]
                               {
                                   "миллионов", "миллион", "миллиона", "миллиона", "миллиона", "миллионов", "миллионов", "миллионов",
                                   "миллионов", "миллионов"
                               };
            var thousands = new[] { "тысяч", "тысяча", "тысячи", "тысячи", "тысячи", "тысяч", "тысяч", "тысяч", "тысяч", "тысяч" };
            var rubles = new[] { "рублей", "рубль", "рубля", "рубля", "рубля", "рублей", "рублей", "рублей", "рублей", "рублей" };

            if (sum <= 0)
            {
                res += "минус ";

                sum = -sum;
            }

            var mainSum = (long)Math.Truncate(sum);
            var subSum = (long)Math.Round((sum - Math.Truncate(sum)) * 100);

            int LDN;
            long i999;

            Math.DivRem(mainSum / 1000000000, 1000, out i999);

            var tmp = SubFormatText((int)i999, 0, out LDN);

            if (!string.IsNullOrEmpty(tmp.TrimEnd(' ')))
                tmp += billions[LDN] + " ";

            res += tmp;

            Math.DivRem(mainSum / 1000000, 1000, out i999);

            tmp = SubFormatText((int)i999, 0, out LDN);

            if (!string.IsNullOrEmpty(tmp.TrimEnd(' ')))
                tmp += millions[LDN] + " ";

            res += tmp;

            Math.DivRem(mainSum / 1000, 1000, out i999);

            tmp = SubFormatText((int)i999, 1, out LDN);

            if (!string.IsNullOrEmpty(tmp.TrimEnd(' ')))
                tmp += thousands[LDN] + " ";

            res += tmp;

            Math.DivRem(mainSum, 1000, out i999);

            tmp = SubFormatText((int)i999, 0, out LDN).TrimEnd(' ');

            res += tmp;

            if (isMoney)
            {
                res += " " + rubles[LDN];

                if (subSum != 0)
                {
                    res += " " + subSum.ToString("00") + " копеек";
                }
            }

            return res;
        }

        private string SubFormatText(int i999, int value, out int LDN)
        {
            var hundreds = new[]
                               {
                                   "", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ",
                                   "девятьсот "
                               };
            var tens = new[]
                           {
                               "", "десять ", "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ",
                               "восемьдесят ", "девяносто "
                           };
            var digits = new[,]
                             {
                                 {"", "один ", "два ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять "},
                                 {"", "одна ", "две ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять "}
                             };

            var tens1P = new[]
                             {
                                 "десять ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", "пятнадцать ",
                                 "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать "
                             };

            var HN = 0;
            var TN = 0;
            var DN = 0;

            Math.DivRem(i999 / 100, 10, out HN);
            Math.DivRem(i999 / 10, 10, out TN);
            Math.DivRem(i999, 10, out DN);

            LDN = DN;

            if (TN == 1) LDN = 0;

            return hundreds[HN] + (TN == 1 ? tens1P[DN] : tens[TN]) + (TN != 1 ? digits[value, DN] : string.Empty);
        }
    }
}
