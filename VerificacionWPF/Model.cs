using System;
using System.Collections.Generic;
using LiteDB;

namespace VerificacionComprobantes.Model
{
	public class Person
	{
		public ObjectId Id { get; set; }

		public string Name { get; set; }

		public string Email { get; set; } 

		public List<Operation> Operations { get; set; }

		public Person ()
		{
		}
	}

	public class Operation {
		public ObjectId Id { get; set; }

		public string Entity { get; set; }

		public string OperationNumber { get; set; }

		public string FacturationDate { get; set; }

		public string Total { get; set; }

		public string Voucher { get; set; }

		public Operation() {
		}
	}
}

