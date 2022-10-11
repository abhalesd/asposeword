using Aspose.Words;
using Aspose.Words.Reporting;
using NUnit.Framework;
using System.Data;

namespace DocsExamples.LINQ_Reporting_Engine
{
    internal class Tables : DocsExamplesBase
    {

        [Test]
        public void HausgeldComplexXmlDataSourceWithSchemaLinq()
        {
            // DataSet ds = new DataSet();
            // ds.ReadXml(MyDir + "HausgeldData/hausgeld_dataset.xml");

            // XmlDataSource source = new XmlDataSource(MyDir + "HausgeldData/hausgeld_full.xml", MyDir + "HausgeldData/hausgeld_full_schema.xml");
            XmlDataSource source = new XmlDataSource(MyDir + "HausgeldData/hausgeld_full.xml");

            Document doc = new Document(MyDir + "HausgeldData/Reporting engine template - Hausgeldabrechnung Komplex.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.AllowMissingMembers;
            engine.Options |= ReportBuildOptions.InlineErrorMessages;

            engine.BuildReport(doc, source);
            // engine.BuildReport(doc, ds);

            doc.Save(ArtifactsDir + "ReportingEngine.HausgeldAbrechnungLinqComplex.docx");
        }

        [Test]
        public void HausgeldComplexXmlDataSourceLinq()
        {
            XmlDataSource source = new XmlDataSource(MyDir + "HausgeldData/hgld_all_data.xml");
            DataSet ds = new DataSet();
            ds.ReadXml(MyDir + "HausgeldData/hgld_all_data.xml");


            Document doc = new Document(MyDir + "HausgeldData/hugelddocumentTest.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.AllowMissingMembers;
            engine.Options |= ReportBuildOptions.InlineErrorMessages;

            engine.BuildReport(doc, source);
            // engine.BuildReport(doc, ds);

            doc.Save(ArtifactsDir + "ReportingEngine.HausgeldAbrechnungLinqComplex.docx");
        }

        [Test]
        public void TestInlineErros()
        {
            XmlDataSource source = new XmlDataSource(MyDir + "Mail merge data - ConditionalData.xml");

            Document doc = new Document(MyDir + "Reporting engine template - TestConditionalData.docx");

            ReportingEngine engine = new ReportingEngine();
            // engine.Options |= ReportBuildOptions.AllowMissingMembers;
            engine.Options |= ReportBuildOptions.InlineErrorMessages;

            engine.BuildReport(doc, source);

            doc.Save(ArtifactsDir + "ReportingEngine.TestConditionalData.docx");
        }

        [Test]
        public void HausgeldXmlDataSourceLinq()
        {
            XmlDataSource source = new XmlDataSource(MyDir + "Mail merge data - hausgeld_full.xml");

            Document doc = new Document(MyDir + "Reporting engine template - Hausgeldabrechnung.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.AllowMissingMembers;
            engine.Options |= ReportBuildOptions.InlineErrorMessages;

            engine.BuildReport(doc, source);

            doc.Save(ArtifactsDir + "ReportingEngine.HausgeldAbrechnungLinq.docx");
        }

        [Test]
        public void TestXmlDataSourceLinq()
        {
            //ExStart:HausgeldAbrechnungLinq
            XmlDataSource source = new XmlDataSource(MyDir + "Mail merge data - test.xml");

            Document doc = new Document(MyDir + "Reporting engine template - Hausgeldabrechnung.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, source);

            doc.Save(ArtifactsDir + "ReportingEngine.HausgeldAbrechnungLinq.docx");
            //ExEnd:HausgeldAbrechnungLinq
        }

        [Test]
        public void TestLinq()
        {
            //ExStart:HausgeldAbrechnungLinq
            DataSet ds = new DataSet();
            ds.ReadXml(MyDir + "Mail merge data - test.xml");
            ds.DataSetName = "ds";

            Document doc = new Document(MyDir + "Reporting engine template - Hausgeldabrechnung.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, ds);

            doc.Save(ArtifactsDir + "ReportingEngine.HausgeldAbrechnungLinq.docx");
            //ExEnd:HausgeldAbrechnungLinq
        }

        [Test]
        public void HausgeldAbrechnungLinq()
        {
            //ExStart:HausgeldAbrechnungLinq
            DataSet ds = new DataSet();
            ds.ReadXml(MyDir + "Mail merge data - hausgeld_full.xml");

            Document doc = new Document(MyDir + "Reporting engine template - Hausgeldabrechnung.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, ds, "Data");

            doc.Save(ArtifactsDir + "ReportingEngine.HausgeldAbrechnungLinq.docx");
            //ExEnd:HausgeldAbrechnungLinq
        }

        [Test]
        public void InTableAlternateContent()
        {
            //ExStart:InTableAlternateContent
            Document doc = new Document(MyDir + "Reporting engine template - Total.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetContracts(), "Contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.InTableAlternateContent.docx");
            //ExEnd:InTableAlternateContent
        }

        [Test]
        public void InTableMasterDetail()
        {
            //ExStart:InTableMasterDetail
            Document doc = new Document(MyDir + "Reporting engine template - Nested data table.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.InTableMasterDetail.docx");
            //ExEnd:InTableMasterDetail
        }

        [Test]
        public void InTableWithFilteringGroupingSorting()
        {
            //ExStart:InTableWithFilteringGroupingSorting
            Document doc = new Document(MyDir + "Reporting engine template - Table with filtering.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetContracts(), "contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.InTableWithFilteringGroupingSorting.docx");
            //ExEnd:InTableWithFilteringGroupingSorting
        }
    }
}