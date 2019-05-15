using System;

namespace BasketManagement
{
    class Program
    {
        static void Main(string[] args)
        {
            var MyBasket = new Basket();
            MyBasket.GiveDataToList();
            MyBasket.SaveText();
            MyBasket.SaveToaJason();
            MyBasket.SaveToExcel();
            MyBasket.LoadfromText();
            MyBasket.LoadfromJson();
            MyBasket.LoadfromExcel();


        }
    }
}
