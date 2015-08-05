using System.Collections.Generic;
using System.Linq;
using Converter.Properties;
using ExcelRLibrary;
using PatternsLib;
using REntities.Oktmo;

namespace Formater
{
    public class ObjectOktmoCoposition
    {
        public List<OktmoRowDTO> SubjectOktmoRows { get; private set; }
        public List<OktmoRowDTO> CustomOktmoRows { get; private set; }

        public ObjectOktmoCoposition([NotNull]List<OktmoRowDTO> subjectOktmoRowDTOs)
        {
            this.SubjectOktmoRows = subjectOktmoRowDTOs;
            ResetToSubject();
        }

        public ObjectOktmoCoposition()
        {
            CustomOktmoRows = new List<OktmoRowDTO>();
            SubjectOktmoRows = new List<OktmoRowDTO>();
        }

        public void FixDoubles()
        {
            CustomOktmoRows = CustomOktmoRows.DistinctBy(
                    r => r.Subject.ToLower() + "/" + r.Region.ToLower().Replace("город","").Trim() + "/" + r.Settlement.Name.ToLower() + "/" + r.City.Name.ToLower() + "/" + r.City.Type.ToLower()).ToList();
        }

        public bool SubjectHasEqualNearCity(string name)
        {
            return SubjectOktmoRows.Any(r => (r.City.Name?? "").ToLower().Equals(name.ToLower()));
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
            return CustomOktmoRows.Any(r => (r.Settlement.Name??"").ToLower().Equals(fullName.ToLower()));
        }

        public bool HasEqualNearCity(string city)
        {
            return CustomOktmoRows.Any(r => (r.City.Name??"").ToLower().Equals(city.ToLower()));
        }

        public bool HasEqualCityType(string typeOfNearCity)
        {
            return CustomOktmoRows.Any(r => (r.City.Type??"").ToLower().Equals(typeOfNearCity.ToLower()));
        }

        public bool HasEqualSubject(string subjectName)
        {
            return CustomOktmoRows.Any(r => (r.Subject ?? "").ToLower().Equals(subjectName.ToLower()));
        }

        public void SetSpecifications(ISpecification<OktmoRowDTO> specs)
        {
            CustomOktmoRows = CustomOktmoRows.FindAll(specs.IsSatisfiedBy).Distinct(new OktmoRowDTOEqualityComparer()).ToList();
        }

        public void SetSubjectRows(List<OktmoRowDTO> rows)
        {
            this.SubjectOktmoRows = rows;
        }

        public void ResetToSubject()
        {
            CustomOktmoRows = new List<OktmoRowDTO>(SubjectOktmoRows);
        }

        public bool SubjectFromOktmo(string subjectName)
        {
            return SubjectOktmoRows.Any(r => (r.Subject ?? "").ToLower().Equals(subjectName.ToLower()));
        }
    }
}