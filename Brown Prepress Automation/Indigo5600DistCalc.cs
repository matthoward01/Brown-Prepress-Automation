using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Brown_Prepress_Automation
{
    class Indigo5600DistCalc
    {
        public (List<string>, List<string>) Indigo5600Dist(FormMain mainForm, List<string> art, List<int> qty, int nUp)
        {
            List<string> item = new List<string>();
            List<int> itemQty = new List<int>();
            List<string> itemPrint = new List<string>();
            List<int> itemQtyPrint = new List<int>();
            List<string> itemHold = new List<string>();
            List<int> itemQtyHold = new List<int>();
            List<string> itemTotal = new List<string>();
            List<string> diffPerPage = new List<string>();
            List<int> nUpfactors = Factor(nUp);            
            nUpfactors.Sort();
            nUpfactors.Reverse();
            List<int> nUpfactorsHold = nUpfactors;
            item = art;
            itemQty = qty;
            int printed = 0;
            int qtyCheck = 0;
            while (item.Count > 0)
            {
                int itemListCount = 0;
                itemHold.Clear();
                itemQtyHold.Clear();
                qtyCheck = itemQty[0];

                foreach (int iqty in itemQty)
                {
                    if (iqty != qtyCheck)
                    {
                        itemHold.Add(item[itemListCount]);
                        itemQtyHold.Add(iqty);
                    }
                    else
                    {
                        itemPrint.Add(item[itemListCount]);
                        itemQtyPrint.Add(iqty);
                    }
                    itemListCount++;
                }
                item = itemHold.ToList();
                itemQty = itemQtyHold.ToList();

                while (itemPrint.Count > 0)
                {
                    nUpfactors = nUpfactorsHold;
                    foreach (int fact in nUpfactors)
                    {
                        while ((itemPrint.Count - fact >= 0) && itemPrint.Count > 0)
                        //if (itemPrint.Count() % fact == 0 && (itemPrint.Count > 0))
                        {
                            int addPage = 0;
                            int divFactor = nUp / fact;
                            for (int i = 0; i < fact; i++)
                            {                                
                                if (fact != nUp)
                                {
                                    for (int z = 0; z < divFactor; z++)
                                    {
                                        itemTotal.Add(itemPrint[addPage]);
                                    }
                                    addPage++;
                                }
                                else
                                {
                                    itemTotal.Add(itemPrint[addPage]);
                                    addPage++;
                                }
                                
                            }

                            itemPrint.RemoveRange(0, fact);
                            printed = (int)Math.Ceiling((double)itemQtyPrint[0] / divFactor);
                            diffPerPage.Add(fact + " Diff - Print " + printed + " Sheets - For Qty of " + itemQtyPrint[0]);
                            itemQtyPrint.RemoveRange(0, fact);
                        }
                    }
                }
            }
            return (itemTotal, diffPerPage);
        }

        public List<int> Factor(int number)
        {
            var factors = new List<int>();
            int max = (int)Math.Sqrt(number);  // Round down

            for (int factor = 1; factor <= max; ++factor) // Test from 1 to the square root, or the int below it, inclusive.
            {
                if (number % factor == 0)
                {
                    factors.Add(factor);
                    if (factor != number / factor) // Don't add the square root twice!  Thanks Jon
                        factors.Add(number / factor);
                }
            }
            return factors;
        }
    }
}
