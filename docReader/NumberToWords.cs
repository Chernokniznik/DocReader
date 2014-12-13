using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace docReader
{
    public  class NumberToWords
    {
        public static string NumberToWord(string rawnumber)
        {
            int inputNum = 0;
            int dig1, dig2, dig3, level = 0, lasttwo, threeDigits, onecent = 0;
            string dollars, cents;

            try
            {
                string[] Splits = new string[2];
                Splits = rawnumber.Split('.');   //notice that it is ' and not "
                inputNum = Convert.ToInt32(Splits[0]);
                dollars = "";
                cents = Splits[1];
                int.TryParse(cents, out onecent);
                if (cents.Length == 1)
                {
                   // cents += "0 ";   // 12.5 is twelve and 50/100, not twelve and 5/100
                }
            }
            catch
            {
                cents = "00";
                inputNum = Convert.ToInt32(rawnumber);
                dollars = "";
            }
            string x = "";
            //they had zero for ones and tens but that gave ninety zero for 90
            string[] ones = { "", "viens", "divi", "trīs", "četri", "pieci", "seši", "septiņi", "astoņi", "deviņi", "desmit", "vienpadsmit", "divpadsmit", "trīspadsmit", "četrpadsmit", "piecpadsmit", "sešpadsmit", "septiņpadsmit", "astoņpadsmit", "deviņpadsmit" };
            string[] tens = { "", "desmit", "divdesmit", "trīsdesmit", "četrdesmit", "piecdesmit", "sešdesmit", "septiņdesmit", "astoņdesmit", "deviņdesmit" };
            string[] thou = { "", "tūkstotis", "milions", "miljards", "trilion", "quadrillion", "quintillion" };

            bool isNegative = false;
            if (inputNum < 0)
            {
                isNegative = true;
                inputNum *= -1;
            }
            if (inputNum == 0)
            {

                return "nulle " + "eiro un " + cents + " centi ";
            }

            string s = inputNum.ToString();
            while (s.Length > 0)
            {
                //Get the three rightmost characters
                x = (s.Length < 3) ? s : s.Substring(s.Length - 3, 3);

                // Separate the three digits
                threeDigits = int.Parse(x);
                lasttwo = threeDigits % 100;
                dig1 = threeDigits / 100;
                dig2 = lasttwo / 10;
                dig3 = (threeDigits % 10);

                // append a "thousand" where appropriate
                if (level > 0 && dig1 + dig2 + dig3 > 0)
                {
                    dollars = thou[level] + " " + dollars;
                    dollars = dollars.Trim();
                }

                // check that the last two digits is not a zero
                if (lasttwo > 0)
                {
                    if (lasttwo < 20)
                    {
                        // if less than 20, use "ones" only
                        dollars = ones[lasttwo] + " " + dollars;
                    }
                    else
                    {
                        // otherwise, use both "tens" and "ones" array
                        dollars = tens[dig2] + " " + ones[dig3] + " " + dollars;
                    }
                    if (s.Length < 3)
                    {

                        if (isNegative)
                        { dollars = "negatīvs " + dollars; }
                        if (onecent == 1)
                        {
                            return dollars + "eiro" + " un " + onecent + " cents";
                        }
                        else
                        {
                            return dollars + "eiro" + " un " + cents + " centi";
                        }
                    }
                }

                // if a hundreds part is there, translate it
                if (dig1 > 0)
                {
                    dollars = ones[dig1] + " simts " + dollars;
                    s = (s.Length - 3) > 0 ? s.Substring(0, s.Length - 3) : "";
                    level++;
                }
                else
                {
                    if (s.Length > 3)
                    {
                        s = s.Substring(0, s.Length - 3);
                        level++;
                    }
                }
            }

            if (isNegative) { dollars = "negatīvs " + dollars; }
            return dollars + "eiro" + " un " + cents + " centi";
        }
    }
}
