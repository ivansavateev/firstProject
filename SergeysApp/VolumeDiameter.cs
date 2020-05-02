using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SergeysApp
{
    class VolumeDiameter : IEquatable<VolumeDiameter>
    {
        public string Volume;
        public string Diameter;
        public string Length;
        public string Weight;
        public string Height;
        public string H2;
        public string H3;
        public string H4;
        public string D2;

        public VolumeDiameter(string vol, string diam)
        {
            Volume = vol;
            Diameter = diam;
        }

        public static bool operator ==(VolumeDiameter obj1, VolumeDiameter obj2)
        {
            if ((obj1.Volume.Equals(obj2.Volume)) && (obj1.Diameter.Equals(obj2.Diameter)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool operator !=(VolumeDiameter obj1, VolumeDiameter obj2)
        {
            if ((obj1.Volume.Equals(obj2.Volume)) && (obj1.Diameter.Equals(obj2.Diameter)))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public bool Equals(VolumeDiameter other)
        {
            return EqualityComparer<string>.Default.Equals(Volume, other.Volume)
                && EqualityComparer<string>.Default.Equals(Diameter, other.Diameter);
        }

        public override bool Equals(object other)
        {
            if (other is VolumeDiameter)
                return Equals((VolumeDiameter)other);

            return false;
        }

        public override int GetHashCode()
        {
            var result = 17;
            unchecked
            {
                result = 31 * result + EqualityComparer<string>.Default.GetHashCode(Volume);
                result = 31 * result + EqualityComparer<string>.Default.GetHashCode(Diameter);
            }
            return result;

        }


    }
}
