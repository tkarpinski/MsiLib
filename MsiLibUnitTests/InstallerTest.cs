using MsiLib;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace MsiInteropUnitTests
{


	/// <summary>
	///This is a test class for InstallerTest and is intended
	///to contain all InstallerTest Unit Tests
	///</summary>
	[TestClass()]
	public class InstallerTest
	{
		/// <summary>
		/// Path to the testing MSI
		/// </summary>
		string path = System.IO.Directory.GetParent(System.IO.Directory.GetCurrentDirectory()).Parent.Parent.FullName + "\\MsiLibUnitTests\\Resources\\TestMsi.Msi"; // TODO: Initialize to an appropriate value

		private TestContext testContextInstance;

		/// <summary>
		///Gets or sets the test context which provides
		///information about and functionality for the current test run.
		///</summary>
		public TestContext TestContext
		{
			get
			{
				return testContextInstance;
			}
			set
			{
				testContextInstance = value;
			}
		}

		#region Additional test attributes
		// 
		//You can use the following additional attributes as you write your tests:
		//
		//Use ClassInitialize to run code before running the first test in the class
		//[ClassInitialize()]
		//public static void MyClassInitialize(TestContext testContext)
		//{
		//}
		//
		//Use ClassCleanup to run code after all tests in a class have run
		//[ClassCleanup()]
		//public static void MyClassCleanup()
		//{
		//}
		//
		//Use TestInitialize to run code before running each test
		//[TestInitialize()]
		//public void MyTestInitialize()
		//{
		//}
		//
		//Use TestCleanup to run code after each test has run
		//[TestCleanup()]
		//public void MyTestCleanup()
		//{
		//}
		//
		#endregion


		/// <summary>
		///A test for Open
		///</summary>
		[TestMethod()]
		public void OpenTest()
		{
			try
			{
				var installer = Installer.Open(path, 1);
				var name = installer.GetProperty("ProductName");
				Assert.AreEqual("TestMsi", name);
			}
			catch (Exception ex)
			{
				Assert.Fail(ex.ToString());
			}
		}

		[TestMethod()]
		public void UpdatePropertyTest()
		{
			try
			{
				var installer = Installer.Open(path, 1);

				var name = installer.GetProperty("ProductName");
				Assert.AreEqual("TestMsi", name);
				installer.SetProperty("ProductName", "TestMsi2");
				name = installer.GetProperty("ProductName");
				Assert.AreEqual("TestMsi2", name);
				installer.SetProperty("ProductName", "TestMsi");
				name = installer.GetProperty("ProductName");
				Assert.AreEqual("TestMsi", name);
			}
			catch (Exception ex)
			{
				Assert.Fail(ex.ToString());
			}
		}
		[TestMethod()]
		public void CreatePropertyTest()
		{
			try
			{
				var installer = Installer.Open(path, 1);
				installer.SetProperty("TEST_PROP2", "TestMsi");
				var name = installer.GetProperty("TEST_PROP2");
				Assert.AreEqual("TestMsi", name);
			}
			catch (Exception ex)
			{
				Assert.Fail(ex.ToString());
			}
		}
		[TestMethod()]
		public void FindInvalidPropertyStringEmpty()
		{
			try
			{
				var installer = Installer.Open(path, 0);
				var name = installer.GetProperty(Guid.NewGuid().ToString());
				Assert.AreEqual("NOVALUE", name);
			}
			catch (Exception ex)
			{
				Assert.Fail(ex.ToString());
			}
		}
		[TestMethod()]
		public void GetIntValue()
		{
			try
			{
				var installer = Installer.Open(path, 0);
				var name = installer.GetProperty("ARPNOMODIFY");
				Assert.AreEqual("NOVALUE", name);
			}
			catch (Exception ex)
			{
				Assert.Fail(ex.ToString());
			}
		}
		[TestMethod()]
		public void UpdateIntValue()
		{
			try
			{
				var installer = Installer.Open(path, 1);
				installer.SetProperty("ARPNOREMOVE", 1);
				var name = installer.GetProperty("ARPNOREMOVE");
				Assert.AreEqual("1", name);
				installer.SetProperty("ARPNOREMOVE", 0);
			}
			catch (Exception ex)
			{
				Assert.Fail(ex.ToString());
			}
		}
	}
}
