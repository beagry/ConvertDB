using System.Collections.Generic;
using System.Linq;
using Converter.Properties;
using ExcelRLibrary.SupportEntities.Oktmo;
using PatternsLib;

namespace Formater
{
    public class OktmoHelper
    {
        public List<OktmoRow> SubjectOktmoRows { get; private set; }
        public List<OktmoRow> CustomOktmoRows { get; private set; }

        public OktmoHelper([NotNull]List<OktmoRow> subjectOktmoRows)
        {
            this.SubjectOktmoRows = subjectOktmoRows;
            ResetToSubject();
        }

        public OktmoHelper()
        {
            CustomOktmoRows = new List<OktmoRow>();
            SubjectOktmoRows = new List<OktmoRow>();
        }

        public bool SubjectHasEqualNearCity(string name)
        {
            return SubjectOktmoRows.Any(r => (r.NearCity ?? "").ToLower().Equals(name.ToLower()));
        }

        public bool SubjectHasEqualRegion(string name)
        {
            return SubjectOktmoRows.Any(r => (r.Region ?? "").ToLower().Equals(name.ToLower()));
        }

        public bool HasEqualRegion(string fullName)
        {
            return CustomOktmoRows.Any(r => (r.Region ?? "").ToLower().Equals(fullName.ToLower()));
        }

        public bool HasEqualSettlement(string fullName)
        {
            return CustomOktmoRows.Any(r => (r.Settlement??"").ToLower().Equals(fullName));
        }

        public bool HasEqualNearCity(string city)
        {
            return CustomOktmoRows.Any(r => (r.NearCity??"").ToLower().Equals(city.ToLower()));
        }

        public bool HasEqualCityType(string typeOfNearCity)
        {
            return CustomOktmoRows.Any(r => (r.TypeOfNearCity??"").ToLower().Equals(typeOfNearCity.ToLower()));
        }

        public bool HasEqualSubject(string subjectName)
        {
            return CustomOktmoRows.Any(r => (r.Subject ?? "").ToLower().Equals(subjectName.ToLower()));
        }

        public void SetSpecifications(ISpecification<OktmoRow> specs)
        {
            CustomOktmoRows = CustomOktmoRows.FindAll(specs.IsSatisfiedBy).ToList();
        }

        public void SetSubjectRows(List<OktmoRow> rows)
        {
            this.SubjectOktmoRows = rows;
        }

        public void ResetToSubject()
        {
            CustomOktmoRows = new List<OktmoRow>(SubjectOktmoRows);
        }

        public bool SubjectFromOktmo(string subjectName)
        {
            return SubjectOktmoRows.Any(r => (r.Subject ?? "").ToLower().Equals(subjectName.ToLower()));
        }
    }
}