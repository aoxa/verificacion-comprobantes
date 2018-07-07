using System;
using System.Collections.Generic;
using System.IO;
using System.Configuration;
using System.Collections.Specialized;
using System.Net.Mail;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using LiteDB;

namespace VerificacionComprobantes.Service
{

	public class ConfigurationServer {
		private static ConfigurationServer instance;
		private static readonly object padlock = new object ();
		public string dbConfig { set; get;}

		private ConfigurationServer() {}

		public static ConfigurationServer Instance {
			get {
				lock(padlock) {
					if (null == instance) {
						instance = new ConfigurationServer();
					}

					return instance; 
				}
			}
		}

		public Configuration getConfiguration() {
			Configuration config;

			using (var db = new LiteDatabase (@dbConfig)) {
				config = GetConfig (db.GetCollection<Configuration> ("config"));
			}

			return config;
		}

		private Configuration GetConfig(LiteCollection<Configuration> collection) {
			IEnumerator<Configuration> enumerator = collection.FindAll ().GetEnumerator ();
			Configuration latest = null;
			while (enumerator.MoveNext()) {
				latest = enumerator.Current;
			}
			return latest;
		}

		public void Store(Configuration config) {
			using (var db = new LiteDatabase (@dbConfig)) {
				LiteCollection<Configuration> collection = db.GetCollection<Configuration> ("config");

				Configuration current = GetConfig(collection);
				if (null != current) {
					current.Email.Server = config.Email.Server;
					current.Email.Password = config.Email.Password;
					current.Email.Header = config.Email.Header;
					current.Email.Port = config.Email.Port;
					current.Email.User = config.Email.User;
					current.Email.Template = config.Email.Template;
					current.Email.CC = config.Email.CC;
					current.Email.Sender = config.Email.Sender;

					current.Sheet.Sheets = config.Sheet.Sheets;
					current.Sheet.Name = config.Sheet.Name;
					current.Sheet.Date = config.Sheet.Date;
					current.Sheet.Operation = config.Sheet.Operation;
					current.Sheet.Voucher = config.Sheet.Voucher;
					current.Sheet.Total = config.Sheet.Total;

					collection.Update (current);
				} else {
					collection.Insert (config);
				}
			}
		}

		public class Configuration {
			public int Id  { get; set; }
			public EmailServerConfig Email { get; set; }
			public SheetConfig Sheet { get; set; }

			public Configuration() {
				Email = new EmailServerConfig();
				Sheet = new SheetConfig();
			}

			public class EmailServerConfig {
				public string Server { get; set; }
				public string Port  { get; set; }
				public string User  { get; set; }
				public string Password  { get; set; }
				public string Header  { get; set; }
				public string Template { get; set; }
				public string CC { get; set; }
				public string Sender { get; set; }
			}

			public  class SheetConfig {
				public string Sheets { get; set; }
				public string Date { get; set; }
				public string Operation { get; set; }
				public string Voucher { get; set; }
				public string Name { get; set; }
				public string Total { get; set; }
			}
		}
	}

	public class EmailServer
	{
		private SmtpClient smtp;

		public EmailServer() {
		}

		private void Init() {
			if (null != smtp) {
				return;
			}

			ConfigurationServer.Configuration config = ConfigurationServer.Instance.getConfiguration ();

			if (string.IsNullOrEmpty (config.Email.Server)) {
				return;
			}
			smtp = new SmtpClient {
				Host = config.Email.Server, 
				Port = int.Parse(config.Email.Port),
				DeliveryMethod = SmtpDeliveryMethod.Network,
				EnableSsl = true,
				UseDefaultCredentials = false,
				Credentials = new System.Net.NetworkCredential (config.Email.User, config.Email.Password)
			};

			ServicePointManager.ServerCertificateValidationCallback = 
				delegate(object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) 
			{ return true; };
		}	
		public void Send(string to, string title, string body) { 
			Send (to, null, title, body);
		}
		public void Send(string to, string cc, string title, string body) {
			Init ();

			MailMessage msg = new MailMessage ("test@test.com", to, title, body);

			if (!string.IsNullOrEmpty (cc)) {
				System.Collections.IEnumerator enumerator = cc.Split (',').GetEnumerator();

				while (enumerator.MoveNext ()) {
					msg.CC.Add (new MailAddress ((string)enumerator.Current));
				}
			}

			msg.IsBodyHtml = true;
			smtp.Send (msg);
		}
	}

	public class StorageService<T> {
		public StorageService() {}

		public IEnumerable<T> Query (Query query, string space) {
			using (var db = new LiteDatabase (@ConfigurationServer.Instance.dbConfig)) {
				LiteCollection<T> collection = db.GetCollection<T> (space);
                
				return collection.Find (query);
			}
		}

		public IEnumerable<T> QueryAll (string space) {
			using (var db = new LiteDatabase (@ConfigurationServer.Instance.dbConfig)) {
				LiteCollection<T> collection = db.GetCollection<T> (space);

				return collection.FindAll ();
			}
		}

        public void Update(Model.Person person)
        {
            using (var db = new LiteDatabase(@ConfigurationServer.Instance.dbConfig))
            {
                var collection = db.GetCollection<Model.Person>("personas");

                collection.Update(person);
            }
        }
    }
}

