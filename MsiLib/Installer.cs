using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MsiLib
{
	/// <summary>
	/// A C# Wrapper around the WindowsInstaller.Installer COM Object
	/// </summary>
	public class Installer
	{

		private string MSIPath;
		dynamic installer = Activator.CreateInstance(Type.GetTypeFromProgID("WindowsInstaller.Installer"));
		dynamic database;
		/// <summary>
		/// 
		/// </summary>
		/// <param name="path"></param>
		/// <param name="OpenStyle"> 0 - ReadOnly, 1 - ReadWrite</param>
		/// <returns></returns>
		public static Installer Open(string path,int OpenStyle)
		{
			var item = new Installer();
			item.MSIPath = path;
			item.database = item.installer.OpenDatabase(path, OpenStyle);
			dynamic components = item.installer.Components;
			//uint i = Msi.MsiOpenPackage(path, ref item.msiHandle);
			return item;
		}
		/// <summary>
		/// Save the MSI back to the same path
		/// </summary>
		public void Save()
		{
			database.Commit();
		}
		/// <summary>
		/// Get a string value for a specified property
		/// </summary>
		/// <param name="propertyName"></param>
		/// <returns>"NOVALUE" if property not found</returns>
		public string GetProperty(string propertyName)
		{
			string retVal = "NOVALUE";
			dynamic view = database.OpenView(String.Format("SELECT Value from Property where Property = '{0}'",propertyName));
			view.Execute();
			dynamic data = view.Fetch();
			if (data != null)
				retVal = data.StringData(1);
			view.Close();
			return retVal;
		}
		/// <summary>
		/// Create or Update the Value of the specified property
		/// </summary>
		/// <param name="propertyName"></param>
		/// <param name="value"></param>
		/// <returns></returns>
		public bool SetProperty(string propertyName,string value)
		{
			dynamic view = database.OpenView(String.Format("SELECT Value from Property where Property = '{0}'", propertyName));
			view.Execute();
			dynamic rowData = view.Fetch();
			if (rowData != null)
			{
				view = database.OpenView("UPDATE Property SET Value = '" + value + "' WHERE Property = '" + propertyName + "'");
				view.Execute();
				view.Close();
			}
			else
			{
				view = database.OpenView(String.Format("INSERT INTO Property (Property, Value) VALUES ('{0}', '{1}')", propertyName, value));
				view.Execute();
				view.Close();
			}
			return true;
		}
		public bool SetProperty(string propertyName, int value)
		{
			dynamic view = database.OpenView(String.Format("SELECT Value from Property where Property = '{0}'", propertyName));
			view.Execute();
			dynamic rowData = view.Fetch();
			if (rowData != null)
			{
				view = database.OpenView("UPDATE Property SET Value = '" + value + "' WHERE Property = '" + propertyName + "'");
				view.Execute();
				view.Close();
			}
			else
			{
				view = database.OpenView(String.Format("INSERT INTO Property (Property, Value) VALUES ('{0}', {1})", propertyName, value));
				view.Execute();
				view.Close();
			}
			return true;
		}
	}
}
