using System;
using System.Collections.Generic;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using LiteDB;

namespace VerificacionComprobantes
{
	namespace Exceptions {
        [Serializable]
		public class SheetNotPresentException : Exception {
			public SheetNotPresentException(string message) : base(message) {
				
			}
		}
	}
	
	public sealed class XLSDataSource
	{
		public class ColumnMappings
		{
			public static string FECHA = "fecha";
			public static string OPERATION = "n operación";
			public static string NAME = "nombre"; 
			public static string MAIL = "mail";
			public static string VOUCHER = "comprobante";
			public static string TOTAL = "total";
			public static string SHEETS = "Medicons,Maternin";
		}

		private static XLSDataSource instance = null;
		private static readonly object padlock = new object ();

		XLSDataSource ()
		{
		}

		public static XLSDataSource Instance {
			get {
				lock (padlock) {
					if (null == instance) {
						instance = new XLSDataSource ();
					}

					return instance; 
				}
			}
		}

		private string getCellValue (ICell cel)
		{
			if (null == cel)
				return null;
			switch (cel.CellType) {
			case CellType.Numeric:
				if (DateUtil.IsCellDateFormatted (cel)) {
					DateTime date = cel.DateCellValue;
					ICellStyle style = cel.CellStyle;
					// Excel uses lowercase m for month whereas .Net uses uppercase
					string format = style.GetDataFormatString ().Replace ('m', 'M');
					return date.ToString (format);
				} else {
					return cel.NumericCellValue.ToString ();
				}
			case CellType.String:
				return cel.StringCellValue.ToLower ().Trim ();
			default:				
				Console.WriteLine ("CellType not handled");
				return null;
			}
		}

        public string LoadEmails(string filename) {
            VerificacionComprobantes.Service.ConfigurationServer configServer = VerificacionComprobantes.Service.ConfigurationServer.Instance;
			string notFound = "";

			using (var db = new LiteDatabase (@configServer.dbConfig)) {
				var collection = db.GetCollection<Model.Person> ("personas");

				IWorkbook workbook = 
					new XSSFWorkbook (File.Open (filename, System.IO.FileMode.Open));
				var sheet = workbook.GetSheetAt (0);
				Dictionary<string, int> columnIndex = null;

				foreach(IRow row in sheet) {
					if (columnIndex == null) {
						int column = 0;
						columnIndex = new Dictionary<string, int> ();
						string name = null;

						while ((name = getCellValue (row.GetCell (column))) != null) {
							columnIndex.Add (name, column++);
						}
					} else {
						string name = TreatName (getCellValue (row.GetCell (columnIndex [ColumnMappings.NAME])));

						if (string.IsNullOrEmpty (name))
							continue;
						string mail = getCellValue (row.GetCell (columnIndex [ColumnMappings.MAIL]));
						Console.Out.WriteLine (mail + " : " + name);

						Model.Person found = collection.FindOne (Query.EQ ("Name", name)); 

						if (null != found) {
							found.Email = mail;

							collection.Update (found);
						} else {
							notFound += name + "\n";
						}
					}
				}
			}

			return notFound;
		}
        public void Init(string filename)
        {
            Init(filename, null);
        }

        public void Init (string filename, System.ComponentModel.BackgroundWorker worker)
		{
			Service.ConfigurationServer configServer = Service.ConfigurationServer.Instance;

			using (var db = new LiteDatabase (@configServer.dbConfig)) {

				var collection = db.GetCollection<Model.Person> ("personas");
				collection.EnsureIndex (x => x.Name);
				collection.EnsureIndex ("Operations.OperationNumber");

				IWorkbook workbook = 
					new XSSFWorkbook (File.Open (filename, System.IO.FileMode.Open));

				int sheets = workbook.NumberOfSheets;

                string[] names = ColumnMappings.SHEETS.Split (',');

                int rowCount = 0;

                if (null != worker)
                {
                    for (int i = 0; i < names.Length; i++)
                    {
                        ISheet sheet = workbook.GetSheet(names[i]);
                        rowCount += sheet.LastRowNum;
                    }
                }   

                int currentRow = 0;

                for (int i = 0; i < names.Length; i++) {
					ISheet sheet = workbook.GetSheet (names [i]);
					if (null == sheet)
						throw new Exceptions.SheetNotPresentException (names [i] + " does not exist");
					Dictionary<string, int> columnIndex = null;

					int column = 0;
                    
					foreach (IRow row in sheet) {

                        if(null != worker) worker.ReportProgress((++currentRow * 100) / rowCount); 


                        if (columnIndex == null) {
							columnIndex = new Dictionary<string, int> ();
							string name = null;

							while ((name = getCellValue (row.GetCell (column))) != null) {
								columnIndex.Add (name, column++);
							}
						} else {		
							string name = TreatName (getCellValue (row.GetCell (columnIndex [ColumnMappings.NAME])));

							if (string.IsNullOrEmpty (name))
								continue;

							Model.Person found = collection.FindOne (Query.EQ ("Name", name)); 

							if (null == found) {
								Model.Person person = BuildPerson (row, columnIndex, names [i]);
								collection.Insert (person);
							} else {

								string operationNumber = getCellValue (row.GetCell (columnIndex [ColumnMappings.OPERATION]));

								Predicate<Model.Operation> isOperation = delegate(Model.Operation op) {
									return op.OperationNumber.Equals(operationNumber);
								};

								Model.Operation operation = found.Operations.Find (isOperation);

								if (null == operation) {
									operation = BuildOperation (row, columnIndex, names [i]); 

									if (null != operation) { 
										found.Operations.Add (operation);

										collection.Update (found);
									}
								} else {									
									string voucher = getCellValue (row.GetCell (columnIndex [ColumnMappings.VOUCHER]));

									if( string.IsNullOrEmpty(operation.Voucher) && 
										!string.IsNullOrEmpty(voucher)) {
										operation.Voucher = voucher;

										collection.Update (found);
									}
								}
							}
						}
					}
				}
			}
		}

		private string TreatName (string name)
		{
			if (null == name)
				return null;
			return name.Replace ("  ", " ").Replace (",", "");
		}

		private Model.Operation BuildOperation (IRow row, Dictionary<string, int> mapping, string entity)
		{
			string fechaCel = getCellValue (row.GetCell (mapping [ColumnMappings.FECHA]));
			string operacionCel = getCellValue (row.GetCell (mapping [ColumnMappings.OPERATION]));
			if (null == operacionCel)
				return null;

			string total = getCellValue (row.GetCell (mapping [ColumnMappings.TOTAL]));

			Model.Operation opeation = new Model.Operation () {
				OperationNumber = operacionCel, Entity = entity,
				FacturationDate = fechaCel, Total = total
			};
				
			string comprobanteCel = getCellValue (row.GetCell (mapping [ColumnMappings.VOUCHER]));

			if (null != comprobanteCel)
				opeation.Voucher = comprobanteCel;

			return opeation;			
		}

		private Model.Person BuildPerson (IRow row, Dictionary<string, int> mapping, string entity)
		{
			string nombreCel = TreatName (getCellValue (row.GetCell (mapping [ColumnMappings.NAME])));

			Model.Person person = new Model.Person () { 
				Operations = new List<Model.Operation> (), Name = nombreCel
			};

			Model.Operation operation = BuildOperation (row, mapping, entity);

			if (null != operation)
				person.Operations.Add (operation);

			return person;
		}
	}
}

