using System.Collections.Generic;
using System.IO;
using System.Linq;
using CsvHelper;
using Tekla.Structures.Model;

namespace BeamReport
{
	public class Program
	{
		static void Main(string[] args)
		{
			ModelObjectEnumerator.AutoFetch = true;
			var modelObjectEnumerator =
				new Model().GetModelObjectSelector().GetAllObjects();

			var modelObjects = modelObjectEnumerator.ToList();

			var beams = modelObjects.OfType<Beam>().ToList();

			var beamsReport = beams.Select(b => new
				{
					b.Name,
					b.Profile.ProfileString,
					b.Class,
					b.Finish,
					b.Material.MaterialString
				})
				.ToList();

			var modelName = new Model().GetInfo().ModelName.Split('.')[0];

			var filePath = $"D:\\{modelName}_BeamReport.csv";
			using (var writer = new CsvWriter(new StreamWriter(filePath)))
			{
				writer.WriteRecords(beamsReport);
			}
		}
	}

	public static class TeklaExtensions
	{
		public static List<ModelObject> ToList(this ModelObjectEnumerator enumerator)
		{
			var modelObjects = new List<ModelObject>();
			while (enumerator.MoveNext())
			{
				var modelObject = enumerator.Current;
				if(modelObject == null) continue;
				modelObjects.Add(modelObject);
			}

			return modelObjects;
		}
	}
}
