using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Hello
{
    class ItemsList
    {
        private List<XmlDocument> listDocumentsItemsList = new List<XmlDocument>();
        List<XmlNode> listNodesXml = new List<XmlNode>();
        public List<ItemAttributes> listAttributesItem = new List<ItemAttributes>();

        string errorItemListFile = "";
        bool condRecognizeItemsListFileMiss = false;

        private int countItemsListAddedPages = 0;
        private int countItemsListXmlFiles = 0;
        public bool condMatchedItemListCollection = false;

        public ItemsList(string itemsListFolderPath)
        {
            string[] itemListsFilePaths = Directory.GetFiles(itemsListFolderPath, "*.xml");

            AssertMe.assert(itemListsFilePaths.Length > 0);

            countItemsListXmlFiles = itemListsFilePaths.Length;

            itemListsFilePaths.Reverse();

            foreach (string itemListsFile in itemListsFilePaths)
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(itemListsFile);
                listDocumentsItemsList.Add(doc);
                XmlNodeList nodes = doc.GetElementsByTagName("_ItemsList");
                listNodesXml.Add(nodes[0]);

                ItemAttributes attributeItem = new ItemAttributes();

                // ДОРАБОТКА Возможно стоит включить в страйк номер делкарации, тк он теперь необязательный элемент
                attributeItem.nameNumberItemDeclaration = nodes[0].SelectSingleNode("_DeclarationNumber").InnerText.Trim();

                // узнаем тип документа
                attributeItem.ItemsListType = nodes[0].SelectSingleNode("_ItemsType").InnerText.Trim();


                int strike = 0;
                int itemNumber = -1;
                if (!int.TryParse(nodes[0].SelectSingleNode("_ItemNumber").InnerText.Trim(), out itemNumber))
                    strike++;
                attributeItem.numberItemInList = itemNumber;

                // берем код товара
                attributeItem.nameCodeItemInList = GetStringItemCode(doc.GetElementsByTagName("_ItemsTable"));

                if (string.IsNullOrEmpty(attributeItem.nameCodeItemInList))
                    strike++;

                attributeItem.nodesTableItemsList = doc.GetElementsByTagName("_ItemDescription");

                if (strike == 2)
                {
                    condRecognizeItemsListFileMiss = true;
                    errorItemListFile += " Не удалось обработать файл списка товаров: " + itemListsFile;
                }
                else
                    listAttributesItem.Add(attributeItem);

            }

            listAttributesItem.Reverse();              

        }

        int GetIntItemCode(XmlNodeList tableNodes) 
        {
            List<int> listCodes = new List<int>();
            int itemNumber = -1;
            foreach (XmlNode tableNode in tableNodes)
            {
                
                if (int.TryParse(tableNode.SelectSingleNode("_ItemCode").InnerText.Trim(), out itemNumber))
                    listCodes.Add(itemNumber);
            }

            IEnumerable<int> top1 = listCodes
            .GroupBy(i => i)
            .OrderByDescending(g => g.Count())
            .Take(1)
            .Select(g => g.Key);

            return top1.Count() > 0 ? top1.First() : -1;
        }

        string GetStringItemCode(XmlNodeList tableNodes)
        {
            List<string> listCodes = new List<string>();

            foreach (XmlNode tableNode in tableNodes)
            {
                    listCodes.Add(tableNode.SelectSingleNode("_ItemCode").InnerText.Trim());
            }

            IEnumerable<string> top1 = listCodes
            .GroupBy(i => i)
            .OrderByDescending(g => g.Count())
            .Take(1)
            .Select(g => g.Key);

            if (top1.Count() == 0)
                return "";

            string stringReturn = top1.First();
            if (!Regex.IsMatch(stringReturn, @"^\d+$"))
                stringReturn = "";

            return stringReturn;
        }


        public class ItemAttributes
        {
            public string ItemsListType { get; set; }

            public string nameNumberItemDeclaration { get; set; }
            public int numberItemInList { get; set; }
            public int numberItemCodeInList { get; set; }
            public string nameCodeItemInList { get; set; }
            public XmlNodeList nodesTableItemsList { get; set; }
        };

    }
}
