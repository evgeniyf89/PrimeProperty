using Microsoft.VisualStudio.TestTools.UnitTesting;
using SoftLine.ActionPlugins;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftLine.ActionPlugins.Tests
{
    [TestClass()]
    public class CreatePrintedFormTests
    {
        private readonly CreatePrintedForm _printedForm;
        public CreatePrintedFormTests()
        {
            _printedForm = new CreatePrintedForm();
        }

        [TestMethod()]
        public void RetrivePrintedFormTest()
        {

            _printedForm.RetrivePrintedForm();
        }

        [TestMethod()]
        public void RenderHtmlToPdfTest()
        {
            _printedForm.RenderHtmlToPdf();
        }
    }
}