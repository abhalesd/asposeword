﻿using System;
using System.Drawing;
using System.IO;
using Aspose.Words;

namespace Insert_Hyperlink_in_document
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for an Aspose.Words license file in the local file system and apply it, if it exists.
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                Aspose.Words.License license = new Aspose.Words.License();

                // Use the license from the bin/debug/ Folder.
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Please make sure to visit ");

            // Specify font formatting for the hyperlink.
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;

            // Insert the link.
            builder.InsertHyperlink("Aspose Website", "http://www.aspose.com", false);

            // Revert to default formatting, and carry on writing.
            builder.Font.ClearFormatting();
            builder.Write(" for more information.");

            doc.Save("Insert_Hyperlink_In_Document.docx");

        }
    }
}
