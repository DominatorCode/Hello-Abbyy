using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Hello
{
    class GTD_Parser
    {
        public int DocumentPagesCount { get; set; }

        public string nameNodeAdditionalDocumentDefinition { get; set; }
        public string nameNodeMainDocumentDefinition { get; set; }

        public int ItemsCount { get; set; } // распознанное количество страниц
        int _countItemsCalculate = 1;

        public string DocumentNumber { get; set; }

        private string DeclarationType { get; set; }

        public string AverageDeclarationNumber { get; set; }

        public string MainXMLFilePath { get; set; }

        public string[] xmlFiles;

        public string finalXMLFileExportPath;
        public string finalXMLFileName;
        public string tmpFolderPath;
        public string pathItemsListFolder { get; set; }

        public bool HasAdditionalPages { get; private set; } = false;
        bool pagesCountMatch = false;
        bool itemsCountMatch = false;

        // содержит информацию о количестве товаров на лист 
        // в дополнительных листах
        List<int> _numberItemsInList = new List<int>();

        public decimal TotalCost { get; set; }

        private XmlDocument mainDoc = new XmlDocument();
        private List<XmlDocument> addDocuments = new List<XmlDocument>();

        private string[] tifFilesCollection;       

        public HashSet<string> errorGTDParser = new HashSet<string>();
        public HashSet<string> warningGTDParser = new HashSet<string>();

        HashSet<int> _listCountItemsOnList = new HashSet<int>();

        ItemsList classItemsList = null;
        List <AdditionalDocument> listClassAdditionalDocument = null;

        public GTD_Parser(string folderPath, string declarationType)
        {

            DeclarationType = declarationType;

            xmlFiles = Directory.GetFiles(folderPath, "*.xml");
            tifFilesCollection = Directory.GetFiles(folderPath, "*.tif");
            
            tifFilesCollection.Reverse();
            xmlFiles.Reverse();

            MainXMLFilePath = xmlFiles[0];
            finalXMLFileName = Path.GetFileName(MainXMLFilePath);
            finalXMLFileExportPath = Path.GetDirectoryName(folderPath);
            mainDoc.Load(MainXMLFilePath);
            tmpFolderPath = folderPath;

            nameNodeMainDocumentDefinition = mainDoc.LastChild.FirstChild.FirstChild.Name;

            DocumentNumber = CastDocumentNumber(mainDoc.GetElementsByTagName("_DeclarationNumber").Item(0).InnerText.ToLower());
           
            warningGTDParser.Add("Предупреждения для документа " + this.DeclarationType + " № " + DocumentNumber);

            try
            {
                DocumentPagesCount = int.Parse(mainDoc.GetElementsByTagName("_PagesCount").Item(0).InnerText.Trim(' '));
            }
            catch
            {
                DocumentPagesCount = -1;
            }

            try
            {
                ItemsCount = int.Parse(mainDoc.GetElementsByTagName("_ItemsCount").Item(0).InnerText.Trim(' '));
            }
            catch
            {
                ItemsCount = -1;
            }

            try
            {
                string test = "";

                if (declarationType == "ГДТ")
                {
                    var result = mainDoc.GetElementsByTagName("_TotalCost");
                    if (result.Count > 0)
                        test = result.Item(0).InnerText.Trim(' ');
                    else
                        test = "-1";

                }
                else
                {
                    var result = mainDoc.GetElementsByTagName("_TotalCostNew");
                    if (result.Count > 0)
                        test = result.Item(0).InnerText.Trim(' ');
                    else
                        test = "-1";
                }

                test = test.Replace('.', ',');    
                TotalCost = decimal.Parse(test, NumberStyles.Currency);
            }
            catch
            {
                TotalCost = -1;
            }

            // очистка таблицы от пустых строк

            XmlNode nodeMainDocument = mainDoc.GetElementsByTagName(nameNodeMainDocumentDefinition).Item(0);

            XmlNode nodeTableForClean = nodeMainDocument.SelectSingleNode("_PaymentTable1");
            if (nodeTableForClean != null)
                DeleteEmptyRows(nodeTableForClean);
            else // либо создаем новую таблицу PaymentTable1
            {
                nodeTableForClean = mainDoc.CreateNode(XmlNodeType.Element, "_PaymentTable1", null);
                nodeMainDocument.InsertAfter(nodeTableForClean, nodeMainDocument.LastChild);
            }

        }

        public void LoadAdditionalDocs()
        {
            AssertMe.assert(xmlFiles.Length > 1);
            HasAdditionalPages = true;

            if (DocumentPagesCount == xmlFiles.Length)
                pagesCountMatch = true;
            else
                warningGTDParser.Add("Указанное количество страниц в документе: (" + DocumentPagesCount + ") и распознаваемое количество страниц: (" + xmlFiles.Length + ") не совпадают");

            listClassAdditionalDocument = new List<AdditionalDocument>();

            for (int i = 1; i < xmlFiles.Length; i++)
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(xmlFiles[i]);
                addDocuments.Add(doc);             
            }

            nameNodeAdditionalDocumentDefinition = addDocuments[0].LastChild.FirstChild.FirstChild.Name;

            foreach (XmlDocument doc in addDocuments)
            {              
                XmlNodeList itemNodes = doc.GetElementsByTagName("_Items");
                _countItemsCalculate += itemNodes.Count;

                // проверяем, сколько товаров размещено на лист
                // должно быть 3, но бывает, что два
                _numberItemsInList.Add(itemNodes.Count);
            }

            if (pagesCountMatch)
                EnumeratePages();

            GetAddDocumensInfo();            

        }

        private void EnumerateItems()
        {
            int itemCount = 2;
            foreach (XmlDocument doc in addDocuments)
            {
                XmlNodeList itemNodes = doc.GetElementsByTagName("_Items");
                foreach (XmlNode node in itemNodes)
                {
                    node.SelectSingleNode("_ItemNumber").InnerText = itemCount.ToString();
                    itemCount++;
                }
            }
        }

        private void EnumerateItems2()
        {
            foreach (XmlDocument doc in addDocuments)
            {
                XmlNodeList itemNodes = doc.GetElementsByTagName("_Items");

                int[] itemsNumbers = new int[itemNodes.Count];

                for (int i = 0; i < itemNodes.Count; i++)
                {
                    try
                    {
                        int itemCount = Int32.Parse(itemNodes[i].SelectSingleNode("_ItemNumber").InnerText);
                        itemsNumbers[i] = itemCount;
                    }
                    catch
                    {
                        itemsNumbers[i] = -1;
                    }
                }

                for (int i = 0; i < itemsNumbers.Length - 1; i++)
                {
                    if (itemsNumbers[i] == -1)
                    {
                        int counter = i + 1;
                        while (counter < itemsNumbers.Length)
                        {
                            if (itemsNumbers[counter] != -1)
                            {
                                itemsNumbers[i] = itemsNumbers[counter] - (counter - i);
                                itemNodes[i].SelectSingleNode("_ItemNumber").InnerText = itemsNumbers[i].ToString();
                                break;
                            }
                            else
                                counter++;
                        }
                    }
                }

                for (int i = itemsNumbers.Length - 1; i > 0; i--)
                {
                    if (itemsNumbers[i] == -1)
                    {
                        int counter = i - 1;
                        while (counter > -1)
                        {
                            if (itemsNumbers[counter] != -1)
                            {
                                itemsNumbers[i] = itemsNumbers[counter] + (i - counter);
                                itemNodes[i].SelectSingleNode("_ItemNumber").InnerText = itemsNumbers[i].ToString();
                                break;
                            }
                            else
                                counter--;
                        }
                    }
                }

            }
        }

        private void EnumeratePages()
        {

            int itemCount = 2;

            foreach (XmlDocument doc in addDocuments)
            {
                XmlNodeList itemNodes = doc.GetElementsByTagName(nameNodeAdditionalDocumentDefinition);
                itemNodes[0].SelectSingleNode("_PageNumber").InnerText = itemCount.ToString();
                itemCount++;
            }

        }


        // объеденяет данные xml файлов в единый файл
        void CombineFiles()
        {

            foreach (XmlDocument doc in addDocuments.Reverse<XmlDocument>())
            {
                
                XmlNodeList nodes = doc.GetElementsByTagName(nameNodeAdditionalDocumentDefinition);

                
                XmlNodeList mainNodes = mainDoc.GetElementsByTagName(nameNodeMainDocumentDefinition);
                XmlNode importNode = mainDoc.ImportNode(nodes[0], true);

                // копируем аттрибуты
                ImportNodeAttributes(nodes[0], importNode, nameNodeAdditionalDocumentDefinition);

                mainNodes[0].ParentNode.InsertAfter(importNode, mainNodes[0]);
            }

        }

        public void RenameAddItemsTagName()
        {
            foreach (XmlDocument doc in addDocuments)
            {
                
                XmlNodeList nodes = doc.GetElementsByTagName(nameNodeAdditionalDocumentDefinition);
                string replaceText = nodes[0].InnerXml;

                int countItemsOnList = _numberItemsInList[addDocuments.IndexOf(doc)];

                // пронумеровываем ячейки Items
                int tagNumber = 1;
                string oldNodeName = "_Items";

                int count = new Regex(Regex.Escape(oldNodeName + ">")).Matches(replaceText).Count / 2;

                while (tagNumber < count + 1)
                {
                    XmlNode newNode = doc.CreateNode(XmlNodeType.Element, oldNodeName + tagNumber.ToString(), nodes[0].SelectSingleNode(oldNodeName).NamespaceURI);
                    newNode.InnerXml = nodes[0].SelectSingleNode(oldNodeName).InnerXml;
                    nodes[0].ReplaceChild(newNode, nodes[0].SelectSingleNode(oldNodeName));
                    tagNumber++;
                }

                // создаем новые ячейки Items для КДТ,
                // если их нет в документе, для последующего заполнения данными из соответсвующего ГДТ
                if (count == 0)
                {
                    while (tagNumber != countItemsOnList + 1)
                    {
                        XmlNode nodeItemNew = doc.CreateNode(XmlNodeType.Element, oldNodeName + tagNumber.ToString(), null);
                        nodes[0].InsertBefore(nodeItemNew, nodes[0].LastChild);
                        tagNumber++;
                    }
                }

                // пронумеровываем таблицы _PaymentTable13
                tagNumber = 1;
                oldNodeName = "_PaymentTable13";
                count = new Regex(Regex.Escape(oldNodeName)).Matches(replaceText).Count / 2;
                int pos = 1;


                // в зависимости от количестватоваров на страницу
                // переименовываем табличные ячейки  
                XmlNode nodeTableForClean = null;
                try
                {
                    nodeTableForClean = nodes[0].SelectSingleNode("_PaymentTable2");
                }
                catch
                {
                    warningGTDParser.Add("не найдена таблица для второго товара");
                }

                switch (_numberItemsInList[addDocuments.IndexOf(doc)])
                {
                    case 1:
                        if (count == 1)
                            RenameTableNode(doc, nodes[0], 1, oldNodeName);
                        else
                        {
                            XmlNode newNodeTable13 = doc.CreateNode(XmlNodeType.Element, oldNodeName.Substring(0, oldNodeName.Length - 2) + pos.ToString(), null);
                            nodes[0].InsertBefore(newNodeTable13, nodes[0].LastChild);
                        }
                        break;
                    case 2:
                        if (count == 2)
                        {
                            if (nodeTableForClean != null)
                            {
                                if (!nodeTableForClean.HasChildNodes)
                                {
                                    nodes[0].RemoveChild(nodeTableForClean);
                                    RenameTableNode(doc, nodes[0], 1, oldNodeName);
                                    RenameTableNode(doc, nodes[0], 2, oldNodeName);
                                }
                                else
                                {
                                    DeleteEmptyRows(nodeTableForClean);
                                    RenameTableNode(doc, nodes[0], 1, oldNodeName);
                                    nodes[0].RemoveChild(nodes[0].SelectSingleNode(oldNodeName));
                                }
                            }
                            else
                            {
                                RenameTableNode(doc, nodes[0], 1, oldNodeName);
                                RenameTableNode(doc, nodes[0], 2, oldNodeName);
                            }
                                                                                     
                        }
                        else if (count == 1)
                        {
                            RenameTableNode(doc, nodes[0], 1, oldNodeName);
                            if (nodeTableForClean != null)
                            {
                                if (nodeTableForClean.HasChildNodes)
                                    DeleteEmptyRows(nodeTableForClean);
                            }
                                
                        }
                        else
                        {
                            CreateNewNodeTable(doc, nodes[0], oldNodeName, 1);
                            if (nodeTableForClean != null)
                            {
                                if (nodeTableForClean.HasChildNodes)
                                    DeleteEmptyRows(nodeTableForClean);
                            }
                            else
                                CreateNewNodeTable(doc, nodes[0], oldNodeName, 2);
                        }
                        break;
                    case 3:
                        {
                            if (count == 2)
                            {
                                RenameTableNode(doc, nodes[0], 1, oldNodeName);
                                RenameTableNode(doc, nodes[0], 3, oldNodeName);
                            }
                            else if (count == 1)
                            {
                                RenameTableNode(doc, nodes[0], 1, oldNodeName);
                                CreateNewNodeTable(doc, nodes[0], oldNodeName, 3);
                            }
                            else
                            {
                                CreateNewNodeTable(doc, nodes[0], oldNodeName, 1);
                                CreateNewNodeTable(doc, nodes[0], oldNodeName, 3);
                            }

                            if (nodeTableForClean != null)
                            {
                                if (nodeTableForClean.HasChildNodes)
                                    DeleteEmptyRows(nodeTableForClean);
                            }
                            else
                                CreateNewNodeTable(doc, nodes[0], oldNodeName, 2);
                            
                            break;
                        }
                }                                   
               
            }
        }

        void AssembleImages(string paramNameDirectoryExport)
        {
            try
            {
                // If only 1 page was passed, copy directly to output
                if (tifFilesCollection.Length == 1)
                {
                    File.Copy(tifFilesCollection[0], paramNameDirectoryExport + "\\" + Path.GetFileName(tifFilesCollection[0]), true); //finalXMLFileExportPath + "\\" + Path.GetFileName(tifFilesCollection[0])
                    return;
                }

                int pageCount = tifFilesCollection.Length;

                // First page
                Image finalImage = Image.FromFile(tifFilesCollection[0]);
                System.Drawing.Imaging.Encoder encoder = System.Drawing.Imaging.Encoder.SaveFlag;
                System.Drawing.Imaging.Encoder encoderComp = System.Drawing.Imaging.Encoder.Compression;

                ImageCodecInfo encoderInfo = ImageCodecInfo.GetImageEncoders().First(i => i.MimeType == "image/tiff");
                EncoderParameters encoderParameters = new EncoderParameters(2);
                encoderParameters.Param[0] = new EncoderParameter(encoder, (long)EncoderValue.MultiFrame);
                encoderParameters.Param[1] = new EncoderParameter(encoderComp, (long)EncoderValue.CompressionCCITT4);

                finalImage.Save(paramNameDirectoryExport + "\\" + Path.GetFileName(tifFilesCollection[0]), encoderInfo, encoderParameters);

                encoderParameters.Param[0] = new EncoderParameter(encoder, (long)EncoderValue.FrameDimensionPage);
                // All other pages
                for (int i = 1; i < pageCount; i++)
                {
                    Image img = Image.FromFile(tifFilesCollection[i]);
                    finalImage.SaveAdd(img, encoderParameters);
                    img.Dispose();
                }

                // Close out the file
                encoderParameters.Param[0] = new EncoderParameter(encoder, (long)EncoderValue.Flush);
                // Last page
                finalImage.SaveAdd(encoderParameters);
                encoderParameters.Dispose();
                finalImage.Dispose();
            }
            catch (Exception ex)
            {
                warningGTDParser.Add("Ошибка при создании Tif файла для ГДТ: " + ex);
            }
        }

        void GetAddDocumensInfo()
        {
            List<string> listNamesDeclarationNumber = new List<string>();

            foreach (XmlDocument doc in addDocuments.Reverse<XmlDocument>())
            {
                XmlNodeList nodes = doc.GetElementsByTagName(nameNodeAdditionalDocumentDefinition);

                XmlNode docNumberNode = nodes[0].SelectSingleNode("_DeclarationNumber");
                string stringNumber = docNumberNode.InnerText.Trim();
                //stringNumber = Regex.Replace(stringNumber, @"\s+", "");

                XmlNode pageNumberNode = nodes[0].SelectSingleNode("_PageNumber");
                int pageNumber = -1;
                int.TryParse(docNumberNode.InnerText.Trim(), out pageNumber);
                listNamesDeclarationNumber.Add(stringNumber);
                listClassAdditionalDocument.Add(new AdditionalDocument(stringNumber, pageNumber));
            }

            // calc average declaration number
            AverageDeclarationNumber = GetAverageStringValue(listNamesDeclarationNumber);

        }

        string GetAverageStringValue(List<string> paramListValues)
        {
            IEnumerable<string> top1 = paramListValues
            .GroupBy(i => i)
            .OrderByDescending(g => g.Count())
            .Take(1)
            .Select(g => g.Key);

            if (top1.Count() == 0)
                return "";

            return top1.First();
        }


        public bool SaveDocumentToFile(string paramPathExport)
        {
            try
            {
                if (HasAdditionalPages)
                    CombineFiles();               
                mainDoc.Save(paramPathExport);
                AssembleImages(Path.GetDirectoryName(paramPathExport));
            }
            catch (Exception ex)
            {
                errorGTDParser.Add("Error to write a final xml: " + ex);
                throw new Exception(string.Join("; ", errorGTDParser));
            }

            if (pagesCountMatch || itemsCountMatch)
                return true;
            else
            {
                warningGTDParser.Add("Указанное количество товаров в документе: (" + ItemsCount + ") и распознанное количество товаров: (" + _countItemsCalculate + ") не совпадают");
                return false;
            }

        }

        public bool InitItemsList(string paramFolderName)
        {
            // load items list
            AssertMe.assert(!String.IsNullOrEmpty(paramFolderName));

            pathItemsListFolder = tmpFolderPath + "\\" + paramFolderName;

            string[] arrayFilesTiffItemsList = Directory.GetFiles(pathItemsListFolder, "*.tif");

            tifFilesCollection = tifFilesCollection.Union(arrayFilesTiffItemsList).ToArray();

            try
            {
                classItemsList = new ItemsList(pathItemsListFolder);
            }
            catch (Exception e)
            {
                errorGTDParser.Add("Не удалось инициализировать класс ItemsList." + e);
                return false;
            }

            return true;

        }

        public bool AddItemsListData()
        {
            bool condInjectedAll = true;
            List<ItemsList.ItemAttributes> copyListItemsAttributes = classItemsList.listAttributesItem.ToList();

            //Debug.Assert(mainDoc.GetElementsByTagName(nameNodeAdditionalDocumentDefinition).Count == 0);

            // inject data in main page
            if (!AddItemsListDataInDocument(mainDoc, 1, DocumentNumber, copyListItemsAttributes))
                condInjectedAll = false;


            if (addDocuments.Count == 0)
                return condInjectedAll;

            // for additional documents
            foreach (XmlDocument additionalDocument in addDocuments)
            {
                int index = addDocuments.IndexOf(additionalDocument);
                if (!AddItemsListDataInDocument(additionalDocument, listClassAdditionalDocument[index].numberDocumentPage, 
                    listClassAdditionalDocument[index].nameDocumentNumber, copyListItemsAttributes))
                    condInjectedAll = false;
            }

            if (copyListItemsAttributes.Count > 0)
            {
                condInjectedAll = false;
                warningGTDParser.Add("Не все листы списка товаров удалось сопоставить для Декларациии №: " + DocumentNumber);
                foreach (ItemsList.ItemAttributes attribute in copyListItemsAttributes)
                {
                    warningGTDParser.Add("Не обработан лист №: " + (attribute.numberItemInList).ToString() + " ,код товара: " + attribute.nameCodeItemInList);
                }

            }

            // сравниваем распознанное количество товаров и подсчитанное
            if (_countItemsCalculate == ItemsCount)
                itemsCountMatch = true;
            else
            {
                DeleteEmptyItemsInLastPage();
                if (_countItemsCalculate == ItemsCount)
                    itemsCountMatch = true;
                else
                    itemsCountMatch = false;
            }
                

            // пронумеровываем товары
            if (pagesCountMatch || itemsCountMatch)
            {
                EnumerateItems();
            }
            else
                EnumerateItems2();

            return condInjectedAll;
        }



        bool AddItemsListDataInDocument(XmlDocument documentInject, int numberPage, string paramNameNumberDocument, List<ItemsList.ItemAttributes> paramCopyListItemsAttributes)
        {
            AssertMe.assert(classItemsList != null);
           
            XmlNodeList nodes = documentInject.GetElementsByTagName("_Items");

            if (nodes.Count == 1)
                numberPage++;
            else
                numberPage = numberPage + 2;


            string[] nameKeyWords = { "товары", "согласно", "прилагаемому", "списку" };

            foreach (XmlNode node in nodes)
            {
                string nameItem = node.SelectSingleNode("_ItemName").InnerText.ToLower();               

                // ДОРАБОТКА расширить условие, добавив код товара к условию
                if (nameKeyWords.Any(w => nameItem.Contains(w)))
                {
                    int numberItem = -1;
                    string nameCodeItem = "";

                    // пытаемся получить номер товара на странице Декларации
                    if (!int.TryParse(node.SelectSingleNode("_ItemNumber").InnerText, out numberItem))
                    {
                        warningGTDParser.Add("Не удалось извлечь номер для списка товаров из документа " + DeclarationType + ":" +
                            paramNameNumberDocument + " , страница: " + numberPage.ToString());

                    }

                    // берем код товара                     
                    nameCodeItem = node.SelectSingleNode("_ItemCode").InnerText.Trim();
                    if (!Regex.IsMatch(nameCodeItem, @"^\d+$"))
                    {
                        nameCodeItem = "";
                        warningGTDParser.Add("Не удалось извлечь код для списка товаров из документа " + DeclarationType + ":" +
                            paramNameNumberDocument + " , страница: " + numberPage.ToString());                        
                    }

                    if (numberItem == -1 & String.IsNullOrEmpty(nameCodeItem))
                        return false;

                    // ищем во всех листах списка товаров нужные данные
                    foreach (ItemsList.ItemAttributes singleItemList in classItemsList.listAttributesItem)
                    {
                        // сначала сравниваем номера
                        if (String.Compare(CastDocumentNumber(singleItemList.nameNumberItemDeclaration), CastDocumentNumber(paramNameNumberDocument), true) != 0)
                            warningGTDParser.Add("Номера документа из " + this.DeclarationType + " №: " + paramNameNumberDocument + " и списка товаров №:" + singleItemList.nameNumberItemDeclaration + " не совпадают");

                        // если номера товара или код совпадают, делаем вставку данных в Декларацию
                        if (singleItemList.numberItemInList == numberItem || string.Compare(singleItemList.nameCodeItemInList, nameCodeItem, true) == 0)
                        {
                            InjectItemListData(node, singleItemList, documentInject);
                            paramCopyListItemsAttributes.Remove(singleItemList);
                            classItemsList.condMatchedItemListCollection = true;

                            _countItemsCalculate = _countItemsCalculate + singleItemList.nodesTableItemsList.Count - 1;
                        }                      

                    }

                    if (!classItemsList.condMatchedItemListCollection)
                    {
                        warningGTDParser.Add("Не удалось соспоставить данные из списка товаров №: " + numberItem + " для " + DeclarationType +
                            " ,страница №: " + numberPage.ToString() + " и по коду товара: " + nameCodeItem);
                        classItemsList.condMatchedItemListCollection = false;
                    }

                }
            }

            return classItemsList.condMatchedItemListCollection;
        }

        void InjectItemListData(XmlNode paramInjectingNode, ItemsList.ItemAttributes paramItemList, XmlDocument paramDocument)
        {
            XmlNode nodeNumberItem = paramInjectingNode.SelectSingleNode("_ItemNumber");
            paramInjectingNode.RemoveAll();
            paramInjectingNode.InsertAfter(nodeNumberItem, paramInjectingNode.FirstChild);

            foreach (XmlNode newItemNode in paramItemList.nodesTableItemsList)
            {
                XmlNode importNode = paramDocument.ImportNode(newItemNode, true);
                paramInjectingNode.InsertAfter(importNode, paramInjectingNode.LastChild);
            }

            EnumerateItemListNodes(paramInjectingNode, paramDocument);
        }

        void EnumerateItemListNodes(XmlNode paramNodeRename, XmlDocument paramDocumentInject)
        {
            
            int tagNumber = 1;
            string nameNode = "_ItemDescription";
            
            foreach (XmlNode node in paramNodeRename.SelectNodes(nameNode))
            {
                XmlNode newNode = paramDocumentInject.CreateNode(XmlNodeType.Element, nameNode + tagNumber.ToString(), node.ParentNode.SelectSingleNode(nameNode).NamespaceURI);
                newNode.InnerXml = node.ParentNode.SelectSingleNode(nameNode).InnerXml;
                node.ParentNode.ReplaceChild(newNode, node.ParentNode.SelectSingleNode(nameNode));
                tagNumber++;
            }
        }

        void DeleteExtraData()
        {

            XmlDocument lastDocument = addDocuments.First();

        }


        string CastDocumentNumber(string paramDocumentNumber) // возвращает первые две секции номера документа
        {
            string stringDocNumber = "";

            int numberEnd = paramDocumentNumber.IndexOf('/', paramDocumentNumber.IndexOf('/') + 1);
            if (numberEnd != -1)
                stringDocNumber = paramDocumentNumber.Substring(0, numberEnd);
            else
                return paramDocumentNumber;

            return Regex.Replace(stringDocNumber, @"\s+", "");
        }

        public void FillKDTEmptyFields(GTD_Parser GDT, bool paramCondContainAddDocuments)
        {

            Debug.Assert(DeclarationType == "КДТ");
            
            var pathGDTDocument = GDT.finalXMLFileExportPath + "\\" + GDT.finalXMLFileName;
            var pathKDTDocument = GDT.finalXMLFileExportPath + "\\" + this.finalXMLFileName;

            XmlDocument xmlGDT = new XmlDocument();
            XmlDocument xmlKDT = new XmlDocument();

            try
            {
                xmlGDT.Load(pathGDTDocument);
                xmlKDT.Load(pathKDTDocument);
            }
            catch (Exception ex)
            {
                this.errorGTDParser.Add("Ошибка заполнения пустых данных КДТ: Не удалось загрузить один из документов. " + ex.Message);
                return;
            }
            

            // сначала заполняем основной лист КДТ
            XmlNodeList listNodesMainDocument = xmlKDT.DocumentElement.GetElementsByTagName(nameNodeMainDocumentDefinition);
            XmlNodeList listNodesDonor = xmlGDT.DocumentElement.GetElementsByTagName(GDT.nameNodeMainDocumentDefinition);

            // Делаем обход по всем ячейкам и заменяем пустые значениями ячеек из ГДТ
            int i = 0;
            if (listNodesMainDocument.Count > 0) // сначала основная страница
            {
                foreach (XmlNode xmlnode in listNodesMainDocument)
                {
                    RemoveNullChildAndAttibute(xmlnode, listNodesDonor[i], nameNodeMainDocumentDefinition);
                    i++;
                }
            }

            // теперь заполняем пустые ячейки в дополнительных листах, если они есть
            if (paramCondContainAddDocuments)
            {
                // проверяем наличие дополнительных листов ГДТ
                if (!GDT.HasAdditionalPages)
                {
                    warningGTDParser.Add("Невозможно заполнить пустые данные добавочного листа КДТ, отсутсвтует соответствующий лист ГДТ");
                    return;
                }

                Debug.Assert(!String.IsNullOrEmpty(nameNodeAdditionalDocumentDefinition));
                Debug.Assert(!String.IsNullOrEmpty(GDT.nameNodeAdditionalDocumentDefinition));

                XmlNodeList listNodesAdditionalKDT = xmlKDT.DocumentElement.GetElementsByTagName(nameNodeAdditionalDocumentDefinition);
                XmlNodeList listNodesAdditionalGDT = xmlGDT.DocumentElement.GetElementsByTagName(GDT.nameNodeAdditionalDocumentDefinition);

                if (listNodesAdditionalGDT.Count != listNodesAdditionalKDT.Count)
                    warningGTDParser.Add("Количество товаров в дополнительных листах ГДТ и КДТ не совпадают");
                else
                {
                    i = 0;
                    foreach (XmlNode additionalKDT in listNodesAdditionalKDT)
                    {
                        RemoveNullChildAndAttibute(additionalKDT, listNodesAdditionalGDT.Item(i), nameNodeAdditionalDocumentDefinition);

                        i++;
                    }
                }
                
            }
            
            xmlKDT.Save(pathKDTDocument);


        }

        void RemoveNullChildAndAttibute(XmlNode xmlNode, XmlNode documentDonor, string rootKDT)
        {
            if (xmlNode.HasChildNodes)
            {
                for (int xmlNodeCount = xmlNode.ChildNodes.Count - 1; xmlNodeCount >= 0; xmlNodeCount--)
                {
                    RemoveNullChildAndAttibute(xmlNode.ChildNodes[xmlNodeCount], documentDonor, rootKDT);
                }
            }
            else if ((String.IsNullOrEmpty(xmlNode.InnerText) & xmlNode.Name != "_ItemCostNew" & xmlNode.Name != "_ItemCostOld" & xmlNode.Name != "_ItemCost") || 
                (xmlNode.Name == "_ItemCostOld" & xmlNode.ParentNode.Name == "_ItemsTable"))
            {
                if (xmlNode.ParentNode != null)
                {
                    var fullNodePath = GetNodePath(xmlNode, rootKDT);
                    XmlNode nodeDonor = null;
                   
                    try
                    {
                        nodeDonor = documentDonor.SelectSingleNode(fullNodePath);
                        if (nodeDonor != null)
                        {
                            if (!String.IsNullOrEmpty(nodeDonor.InnerText))
                            {
                                XmlNode importNode = xmlNode.OwnerDocument.ImportNode(nodeDonor, true);


                                // копируем аттрибуты ячейки
                                XmlElement elementXmlDonor = documentDonor.SelectSingleNode(fullNodePath) as XmlElement;
                                XmlAttributeCollection collectionAttributesXml;
                                if (elementXmlDonor.HasAttributes)
                                {

                                    XmlElement elementXml = xmlNode as XmlElement;
                                    collectionAttributesXml = elementXmlDonor.Attributes;
                                    foreach (XmlAttribute AttributeXml in collectionAttributesXml)
                                    {
                                        if (String.Compare(AttributeXml.Name, "addData:ErrorRef") == 0)
                                        {

                                            XmlAttribute attribute = xmlNode.OwnerDocument.CreateAttribute(AttributeXml.Name);
                                            attribute.Value = AttributeXml.Value;

                                            xmlNode.Attributes.Append(attribute);
                                        }
                                    }


                                }


                                xmlNode.ParentNode.InsertAfter(importNode, xmlNode);
                                xmlNode.ParentNode.RemoveChild(xmlNode);

                                // ДОРАБОТКА: Убрать все лишние атрибуты
                                //string strXMLPattern = @"xmlns(:\w+)?=""([^""]+)""|xsi(:\w+)?=""([^""]+)""";
                                //xml = Regex.Replace(xml, strXMLPattern, "");

                            }
                        }
                        else
                        {
                            // ДОРАБОТКА: указать точный адрес ячейки
                            this.warningGTDParser.Add("Не удалось заполнить пустое значение ячейки КДТ значением из ГДТ");
                        }
                        
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Ошибка обработки КДТ: Не удалось сделать замену ячейки из ГДТ: " + ex);
                    }

                }
            }
            else if (!String.IsNullOrEmpty(xmlNode.InnerText) & xmlNode.Name == "#text" & xmlNode.ParentNode.Name != "_ItemCostNew" &
                xmlNode.ParentNode.Name != "_PageNumber" & xmlNode.ParentNode.Name != "_ItemCostOld") // добавляем атрибут для подсветки ГДТ значений
            {
                XmlAttribute attribute = xmlNode.OwnerDocument.CreateAttribute("DoHighlight");
                attribute.Prefix = "addData";
                attribute.Value = "Yes";

                xmlNode.ParentNode.Attributes.Append(attribute);
            }
        }

        // Get the node full path
        string GetNodePath(XmlNode node, string stopPath)
        {
            string path = node.Name;
            XmlNode search = null;

            if (node.ParentNode == null)
                return path;

            // Get up until ROOT
            while ((search = node.ParentNode).Name != stopPath)
            {
                path = search.Name + "/" + path; // Add to path
                node = search;

                if (node.ParentNode == null) break;
            }

            return path;
        }

        void ImportNodeAttributes(XmlNode nodeDonor, XmlNode nodeImport, string root)
        {
            if (nodeDonor.HasChildNodes)
            {
                for (int xmlNodeCount = nodeDonor.ChildNodes.Count - 1; xmlNodeCount >= 0; xmlNodeCount--)
                {
                    ImportNodeAttributes(nodeDonor.ChildNodes[xmlNodeCount], nodeImport, root);
                }
            }
            else if (nodeDonor.Name != "#text")
            {
                // копируем аттрибуты ячейки
                XmlElement elementXmlDonor = nodeDonor as XmlElement;
                XmlAttributeCollection collectionAttributesXml;
                if (elementXmlDonor.HasAttributes)
                {

                    XmlElement elementXml = nodeImport.SelectSingleNode(GetNodePath(nodeImport, root)) as XmlElement;
                    collectionAttributesXml = elementXmlDonor.Attributes;
                    foreach (XmlAttribute AttributeXml in collectionAttributesXml)
                    {
                        if (String.Compare(AttributeXml.Name, "addData:ErrorRef") == 0)
                        {

                            XmlAttribute attribute = nodeImport.OwnerDocument.CreateAttribute(AttributeXml.Name);
                            attribute.Value = AttributeXml.Value;

                            nodeImport.Attributes.Append(attribute);
                        }
                    }


                }
            }
            
        }


        XmlNode RemoveAllNamespaces(XmlNode documentElement)
        {
            var xmlnsPattern = "\\s+xmlns\\s*(:\\w)?\\s*=\\s*\\\"(?<url>[^\\\"]*)\\\"";
            var outerXml = documentElement.OuterXml;
            var matchCol = Regex.Matches(outerXml, xmlnsPattern);
            foreach (var match in matchCol)
                outerXml = outerXml.Replace(match.ToString(), "");

            var result = new XmlDocument();
            result.LoadXml(outerXml);

            return result;
        }

        // удаляет пустые строки в таблице
        XmlNode DeleteEmptyRows(XmlNode xmlNodeTable)
        {
            XmlNodeList itemNodes = xmlNodeTable.SelectNodes("_tPayments");

            if (itemNodes.Count == 0)
                return xmlNodeTable;

            for (int i = itemNodes.Count - 1; i >= 0; i--)
            {
                string fieldDescript;
                try
                {
                    fieldDescript = itemNodes[i].SelectSingleNode("_Type").InnerText;
                }
                catch
                {
                    return xmlNodeTable;
                }

                if (String.IsNullOrEmpty(fieldDescript) || fieldDescript.Length < 3)
                {
                    itemNodes[i].ParentNode.RemoveChild(itemNodes[i]);
                }
            }

            return xmlNodeTable;
        }

        void RenameTableNode(XmlDocument doc, XmlNode nodeTable, int pos, string oldNodeName)
        {
            XmlNode newNode = doc.CreateNode(XmlNodeType.Element, oldNodeName.Substring(0, oldNodeName.Length - 2) + pos.ToString(), nodeTable.SelectSingleNode(oldNodeName).NamespaceURI);
            newNode.InnerXml = nodeTable.SelectSingleNode(oldNodeName).InnerXml;
            newNode = DeleteEmptyRows(newNode);
            nodeTable.ReplaceChild(newNode, nodeTable.SelectSingleNode(oldNodeName));
        }

        void CreateNewNodeTable(XmlDocument doc, XmlNode node, string oldNodeName, int pos)
        {
            XmlNode newNodeTable13 = doc.CreateNode(XmlNodeType.Element, oldNodeName.Substring(0, oldNodeName.Length - 2) + pos.ToString(), null);
            node.InsertBefore(newNodeTable13, node.LastChild);
        }

        void DeleteEmptyItemsInLastPage()
        {            
            XmlDocument lastDocument = addDocuments.Last();
            XmlNodeList itemNodes = lastDocument.GetElementsByTagName("_Items");

            if (itemNodes.Count < 2) return;

            XmlNodeList nodesTable13 = lastDocument.GetElementsByTagName("_PaymentTable13");
            XmlNodeList nodesTable2 = lastDocument.GetElementsByTagName("_PaymentTable2");

            for (int i = itemNodes.Count - 1; i > 0 ; i--)
            {
                XmlNode node = itemNodes[i];
                int countNodes = itemNodes.Count;
                string nodeToCheck = node.SelectSingleNode("_ItemNumber").InnerText;
                int numberItem = -1;
                Int32.TryParse(nodeToCheck, out numberItem);

                int countHitsDelete = 0; // счетчик для учета количества пустых полей
                if (numberItem > _countItemsCalculate & numberItem != -1 || numberItem < _countItemsCalculate - 3 & numberItem != -1)
                    countHitsDelete++;

                nodeToCheck = node.SelectSingleNode("_ItemName").InnerText;
                if (!String.IsNullOrEmpty(nodeToCheck))
                {
                    if (nodeToCheck.Length < 4)
                        countHitsDelete++;
                }
                else
                    countHitsDelete++;

                nodeToCheck = node.SelectSingleNode("_ItemCode").InnerText;
                if (!String.IsNullOrEmpty(nodeToCheck))
                {
                    if (nodeToCheck.Length < 4)
                        countHitsDelete++;
                }
                else
                    countHitsDelete++;


                nodeToCheck = node.SelectSingleNode("_ItemPrice").InnerText;
                if (!String.IsNullOrEmpty(nodeToCheck))
                {
                    if (nodeToCheck.Length < 3)
                        countHitsDelete++;
                }
                else
                    countHitsDelete++;


                nodeToCheck = node.SelectSingleNode("_StaticPrice").InnerText;
                if (!String.IsNullOrEmpty(nodeToCheck))
                {
                    if (nodeToCheck.Length < 3)
                        countHitsDelete++;
                }
                else
                    countHitsDelete++;


                // если больше трех пустых полей, то удаляем ячейку
                if (countHitsDelete > 2)
                {
                    node.ParentNode.RemoveChild(node);

                    // так же удаляем таблицы
                    if (countNodes == 3)
                    {
                        nodesTable13[1].ParentNode.RemoveChild(nodesTable13[1]);
                        countNodes--;
                    }
                    else
                    {
                        nodesTable2[0].ParentNode.RemoveChild(nodesTable2[0]);
                        countNodes--;
                    }                        
                }

                _countItemsCalculate--;
            }
        }

    }

}
