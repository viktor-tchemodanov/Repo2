﻿using Aspose.Words.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Signature
{
    class SigningSignatureLine
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithSignature();
            
            if (!File.Exists(dataDir + "signature.pfx"))
            {
                Console.WriteLine("Certificate file does not exist.");
                return;
            }
            SimpleDocumentSigning(dataDir);
            SigningEncryptedDocument(dataDir);
            CreatingAndSigningNewSignatureLine(dataDir);
            SigningExistingSignatureLine(dataDir);
            SetSignatureProviderID(dataDir);
            CreateNewSignatureLineAndSetProviderID(dataDir);
        }

        public static void SimpleDocumentSigning(String dataDir)
        {
            // ExStart:SimpleDocumentSigning
            CertificateHolder certHolder = CertificateHolder.Create(dataDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(dataDir + "Document.Signed.docx", dataDir + "Document.Signed_out.docx", certHolder);

            // ExEnd:SimpleDocumentSigning
            Console.WriteLine("\nDocument is signed successfully.\nFile saved at " + dataDir + "Document.Signed_out.docx");
        }

        public static void SigningEncryptedDocument(String dataDir)
        {
            // ExStart:SigningEncryptedDocument

            SignOptions signOptions = new SignOptions();
            signOptions.DecryptionPassword = "decryptionPassword";

            CertificateHolder certHolder = CertificateHolder.Create(dataDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(dataDir + "Document.Signed.docx", dataDir + "Document.EncryptedDocument_out.docx", certHolder, signOptions);
            // ExEnd:SigningEncryptedDocument
            Console.WriteLine("\nDocument is signed with successfully.\nFile saved at " + dataDir + "Document.EncryptedDocument_out.docx");

        }

        public static void CreatingAndSigningNewSignatureLine(String dataDir)
        {
            // ExStart:CreatingAndSigningNewSignatureLine
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            SignatureLine signatureLine = builder.InsertSignatureLine(new SignatureLineOptions()).SignatureLine;
            doc.Save(dataDir + "Document.NewSignatureLine.docx");

            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = signatureLine.Id;
            signOptions.SignatureLineImage = File.ReadAllBytes(dataDir + "SignatureImage.emf");

            CertificateHolder certHolder = CertificateHolder.Create(dataDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(dataDir + "Document.NewSignatureLine.docx", dataDir + "Document.NewSignatureLine.docx_out.docx", certHolder, signOptions);
            // ExEnd:CreatingAndSigningNewSignatureLine

            Console.WriteLine("\nDocument is created and Signed with new SignatureLine successfully.\nFile saved at " + dataDir + "Document.NewSignatureLine.docx_out.docx");
        }

        public static void SigningExistingSignatureLine(String dataDir)
        {
            // ExStart:SigningExistingSignatureLine
            Document doc = new Document(dataDir + "Document.Signed.docx");
            SignatureLine signatureLine = ((Shape)doc.FirstSection.Body.GetChild(NodeType.Shape, 0, true)).SignatureLine;

            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = signatureLine.Id;
            signOptions.SignatureLineImage = File.ReadAllBytes(dataDir + "SignatureImage.emf");

            CertificateHolder certHolder = CertificateHolder.Create(dataDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(dataDir + "Document.Signed.docx", dataDir + "Document.Signed.ExistingSignatureLine.docx", certHolder, signOptions);
            // ExEnd:SigningExistingSignatureLine

            Console.WriteLine("\nDocument is signed with existing SignatureLine successfully.\nFile saved at " + dataDir + "Document.Signed.ExistingSignatureLine.docx");
        }
        
        public static void SetSignatureProviderID(String dataDir)
        {
            // ExStart:SetSignatureProviderID
            Document doc = new Document(dataDir + "Document.Signed.docx");
            SignatureLine signatureLine = ((Shape)doc.FirstSection.Body.GetChild(NodeType.Shape, 0, true)).SignatureLine;

            //Set signature and signature line provider ID
            SignOptions signOptions = new SignOptions();
            signOptions.ProviderId = signatureLine.ProviderId;
            signOptions.SignatureLineId = signatureLine.Id;

            CertificateHolder certHolder = CertificateHolder.Create(dataDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(dataDir + "Document.Signed.docx", dataDir + "Document.Signed_out.docx", certHolder, signOptions);

            // ExEnd:SetSignatureProviderID

            Console.WriteLine("\nProvider ID of signature is set successfully.\nFile saved at " + dataDir + "Document.Signed_out.docx");
        }

        public static void CreateNewSignatureLineAndSetProviderID(String dataDir)
        {
            // ExStart:CreateNewSignatureLineAndSetProviderID
            Document doc = new Document(dataDir + "Document.Signed.docx");

            DocumentBuilder builder = new DocumentBuilder(doc);
            SignatureLine signatureLine = builder.InsertSignatureLine(new SignatureLineOptions()).SignatureLine;
            signatureLine.ProviderId = new Guid("{F5AC7D23-DA04-45F5-ABCB-38CE7A982553}");
            doc.Save(dataDir + "Document.Signed_out.docx");

            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = signatureLine.Id;
            signOptions.ProviderId = signatureLine.ProviderId;

            CertificateHolder certHolder = CertificateHolder.Create(dataDir + "signature.pfx", "signature");
            DigitalSignatureUtil.Sign(dataDir + "Document.Signed_out.docx", dataDir + "Document.Signed_out.docx", certHolder, signOptions);

            // ExEnd:CreateNewSignatureLineAndSetProviderID

            Console.WriteLine("\nCreate new signature line and set provider ID successfully.\nFile saved at " + dataDir + "Document.Signed_out.docx");
        }

    }
}
