
namespace CsvHelper.Excel.Specs
{
    public class Person
    {
        public string Name { get; set; }
        
        public int Age { get; set; }

        public override int GetHashCode()
        {
            int hash = 17;
            hash = hash * 23 + Name?.GetHashCode() ?? 0;
            hash = hash * 23 + Age.GetHashCode();
            return hash;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Person)) return false;
            var other = obj as Person;
            if (Name != other.Name) return false;
            if (Age != other.Age) return false;
            return true;
        }
    }
}