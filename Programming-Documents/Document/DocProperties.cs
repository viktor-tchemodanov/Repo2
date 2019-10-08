﻿
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Properties;
namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocProperties
    {
        public static void Run()
        {            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Enumerates through all built-in and custom properties in a document.
            EnumerateProperties(dataDir);
            // Checks if a custom property with a given name exists in a document and adds few more custom document properties.
            CustomAdd(dataDir);
            // Removes a custom document property.
            CustomRemove(dataDir);

            RemovePersonalInformation(dataDir);
            ConfiguringLinkToContent(dataDir);
        }
        public static void EnumerateProperties(string dataDir)
        {
            // ExStart:EnumerateProperties            
            string fileName = dataDir + "Properties.doc";
            Document doc = new Document(fileName);
            Console.WriteLine("1. Document name: {0}", fileName);

            Console.WriteLine("2. Built-in Properties");
            foreach (DocumentProperty prop in doc.BuiltInDocumentProperties)
                Console.WriteLine("{0} : {1}", prop.Name, prop.Value);

            Console.WriteLine("3. Custom Properties");
            foreach (DocumentProperty prop in doc.CustomDocumentProperties)
                Console.WriteLine("{0} : {1}", prop.Name, prop.Value);
            // ExEnd:EnumerateProperties
        }
        public static void CustomAdd(string dataDir)
        {
            // ExStart:CustomAdd            
            Document doc = new Document(dataDir + "Properties.doc");
            CustomDocumentProperties props = doc.CustomDocumentProperties;
            if (props["Authorized"] == null)
            {
                props.Add("Authorized", true);
                props.Add("Authorized By", "John Smith");
                props.Add("Authorized Date", DateTime.Today);
                props.Add("Authorized Revision", doc.BuiltInDocumentProperties.RevisionNumber);
                props.Add("Authorized Amount", 123.45);
            }
            // ExEnd:CustomAdd
        }
        public static void CustomRemove(string dataDir)
        {
            // ExStart:CustomRemove            
            Document doc = new Document(dataDir + "Properties.doc");
            doc.CustomDocumentProperties.Remove("Authorized Date");
            // ExEnd:CustomRemove
        }
        public static void RemovePersonalInformation(string dataDir)
        {
            // ExStart:RemovePersonalInformation            
            Document doc = new Document(dataDir + "Properties.doc");
            doc.RemovePersonalInformation = true;

            dataDir = dataDir + "RemovePersonalInformation_out.docx";
            doc.Save(dataDir);
            // ExEnd:RemovePersonalInformation
            Console.WriteLine("\nPersonal information has removed from document successfully.\nFile saved at " + dataDir);
        }
        public static void ConfiguringLinkToContent(string dataDir)
        {
            // ExStart:ConfiguringLinkToContent            
            Document doc = new Document(dataDir + "test.docx");

            // Retrieve a list of all custom document properties from the file.
            CustomDocumentProperties customProperties = doc.CustomDocumentProperties;

            // Add linked to content property.
            DocumentProperty customProperty = customProperties.AddLinkToContent("PropertyName", "BookmarkName");

            // Also, accessing the custom document property can be performed by using the property name.
            customProperty = customProperties["PropertyName"];

            // Check whether the property is linked to content.
            bool isLinkedToContent = customProperty.IsLinkToContent;

            // Get the source of the property.
            string source = customProperty.LinkSource;

            // Get the value of the property.
            string value = customProperty.Value.ToString();
            // ExEnd:ConfiguringLinkToContent
        }
    }
}
