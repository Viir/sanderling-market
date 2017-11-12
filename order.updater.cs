#r "System.Windows.Forms"

using Parse = Sanderling.Parse;
using MemoryStruct = Sanderling.Interface.MemoryStruct;
using System.IO;

//TODO: Add the date of the first buy/sell to the MarketLog.txt file so I can tell how old the order is
//TODO: Add 20/10/2017 to all the blank existing entries.
//Host.Break(); // Halts execution until user continues

//WARNING WARNING WARNING - Filter must be set to station only

var Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
string inputFileName = @"C:\Users\Jason\Documents\Visual Studio 2017\Projects\GitHub Eve\trunk\MarketLog.txt";
string orderFileName = @"C:\Users\Jason\Documents\Visual Studio 2017\Projects\GitHub Eve\trunk\MarketOrders.txt";
bool ContainsBlueBackground(MemoryStruct.IListEntry Entry)=>Entry?.ListBackgroundColor?.Any(BackgroundColor=>111<BackgroundColor?.OMilli && 777<BackgroundColor?.BMilli && BackgroundColor?.RMilli<111 && BackgroundColor?.GMilli<111) ?? false;
bool ContainsGreenBackground(MemoryStruct.IListEntry Entry)=>Entry?.ListBackgroundColor?.Any(BackgroundColor=>111<BackgroundColor?.OMilli && 777<BackgroundColor?.GMilli && BackgroundColor?.RMilli<111 && BackgroundColor?.BMilli<111) ?? false;
bool ContainsBlackBackground(MemoryStruct.IListEntry Entry)=>Entry?.ListBackgroundColor?.Any(BackgroundColor=>BackgroundColor?.OMilli>450 && BackgroundColor?.BMilli>240 && BackgroundColor?.RMilli>240 && BackgroundColor?.GMilli>240) ?? true;
bool MatchingOrder(MemoryStruct.IListEntry Entry)=>Entry?.LabelText?.Any(someText=>someText.Text.ToString().RegexMatchSuccess(orderName)) ?? false;
List<FileOrderEntry>fileOrderEntries = new List<FileOrderEntry>();
Random rnd = new Random();
IWindow ModalUIElement=>Measurement?.EnumerateReferencedUIElementTransitive()?.OfType<IWindow>()?.Where(window=>window?.isModal ?? false)?.OrderByDescending(window=>window?.InTreeIndex ?? int.MinValue)?.FirstOrDefault();

string orderName = "";
double defaultMargin = 15.0;
bool foundNew = false;

//Allow me time to move the mouse after starting app.
Host.Delay(5000);

//Read MarketLog
using(FileStream fileStream = new FileStream(inputFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite)) {
  using(var reader = new StreamReader(fileStream)) {
    while (!reader.EndOfStream) {
      var line = reader.ReadLine();
      var values = line.Split(',');
      try {
        if (!values[0].ToString().Equals("Blank")) {
          double minPrice = CalcMinPrice(Convert.ToDouble(values[3]), Convert.ToDouble(values[5]), defaultMargin);
          double maxPrice = CalcMaxPrice(Convert.ToDouble(values[3]), Convert.ToDouble(values[7]), defaultMargin);

          FileOrderEntry newFileOrder = new FileOrderEntry(values[0].ToString(), values[1].ToString(), Convert.ToDouble(values[3]), minPrice, maxPrice, defaultMargin, 0, 0.00, Convert.ToDateTime(values[9]), Convert.ToDateTime(values[11]), Convert.ToInt32(values[13]), Convert.ToBoolean(values[15]));
          fileOrderEntries.Add(newFileOrder);
        }
      } catch {
        Host.Log("Couldn't read: " + line.ToString());
      }
    }
  }
}

//Read MarketOrders and buy anything there
try {
  using(FileStream fileStream = new FileStream(orderFileName, FileMode.Open, FileAccess.Read)) {
    using(var reader = new StreamReader(fileStream)) {
      while (!reader.EndOfStream) {
        var line = reader.ReadLine();
        var values = Regex.Split(line, @"\t ");

        string itemName = values[0];
        string itemQuantity = values[1];
        string itemPrice = values[2];
        string itemSellPrice = values[3];
        
        itemPrice = itemPrice.Replace(",", "");
        itemQuantity = itemQuantity.Replace(",", "");
        itemQuantity = Convert.ToInt32(Convert.ToDouble(itemQuantity)).ToString();
        itemSellPrice = itemSellPrice.Replace(",", "");
        
        bool foundItem = false;
        foreach(FileOrderEntry fileEntry in fileOrderEntries) {
          if (itemName.Equals(fileEntry.Name) && fileEntry.Type.Equals("Buy Order")) {
            foundItem = true;
          }
        }
        if (foundItem == false) {
          //Click My Orders
          Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.RightTabGroup?.ListTab[2]?.RegionInteraction);

          //Wait for MyOrders to be populated
          Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
          var myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
          while (myOrders == null) {
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.RightTabGroup?.ListTab[2]?.RegionInteraction);
            myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
            Host.Delay(5000);
          }

          Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
          myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
          if (myOrders?.BuyOrderView != null && myOrders?.SellOrderView != null) {
            Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.InputText?.FirstOrDefault()?.RegionInteraction);
            Sanderling.KeyboardPressCombined(new[] {
              VirtualKeyCode.LCONTROL,
              VirtualKeyCode.VK_A
            });
            Sanderling.KeyboardPress(VirtualKeyCode.DELETE);
            Sanderling.TextEntry(itemName);
            Host.Delay(1000);
            Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.ButtonText?.FirstOrDefault()?.RegionInteraction);
            Host.Delay(5000);

            MakeSureDoClick: Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            var ListOfEntries = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.LabelText?.ToArray();
            bool doneClick = false;
            foreach(var entry in ListOfEntries) {
              if (entry.Text.Equals(itemName)) {
                Sanderling.MouseClickLeft(entry.RegionInteraction);
                doneClick = true;
                break;
              }
            }
            if (!doneClick) {
              Host.Delay(1000);
              goto MakeSureDoClick;
            }

            Host.Delay(5000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            foreach(var button in Measurement?.WindowRegionalMarket?.FirstOrDefault()?.ButtonText) {
              if (button.Text.Equals("Export to File")) {
                Sanderling.MouseClickLeft(button.RegionInteraction);
                break;
              }
            }
            Host.Delay(1000);
            foreach(var button in Measurement?.WindowRegionalMarket?.FirstOrDefault()?.ButtonText) {
              if (button.Text.Equals("Place Buy Order")) {
                Sanderling.MouseClickLeft(button.RegionInteraction);
                break;
              }
            }
            var quantity = Measurement?.WindowMarketAction?.FirstOrDefault()?.InputText?.ElementAt(1)?.Text;
            while(quantity == null) {
              Host.Delay(500);
              Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
              quantity = Measurement?.WindowMarketAction?.FirstOrDefault()?.InputText?.ElementAt(1)?.Text;            
            }
            while(!quantity.ToString().Equals("1")) {
              Host.Delay(500);
              Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
              quantity = Measurement?.WindowMarketAction?.FirstOrDefault()?.InputText?.ElementAt(1)?.Text;            
            }
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            Sanderling.MouseClickLeft(Measurement?.WindowMarketAction?.FirstOrDefault()?.InputText?.ElementAt(0).RegionInteraction);
            Host.Delay(500);
            Sanderling.MouseClickLeft(Measurement?.WindowMarketAction?.FirstOrDefault()?.InputText?.ElementAt(0).RegionInteraction);
            Host.Delay(500);
            Sanderling.KeyboardPressCombined(new[] {
              VirtualKeyCode.LCONTROL,
              VirtualKeyCode.VK_A
            });
            Sanderling.KeyboardPress(VirtualKeyCode.DELETE);
            double newPrice = Convert.ToDouble(itemPrice);
            newPrice += 20;
            EnterPrice(newPrice);
            Host.Delay(500);
            
            Sanderling.MouseClickLeft(Measurement?.WindowMarketAction?.FirstOrDefault()?.InputText?.ElementAt(1).RegionInteraction);
            Host.Delay(500);
            Sanderling.MouseClickLeft(Measurement?.WindowMarketAction?.FirstOrDefault()?.InputText?.ElementAt(1).RegionInteraction);
            Host.Delay(500);
            Sanderling.KeyboardPressCombined(new[] {
              VirtualKeyCode.LCONTROL,
              VirtualKeyCode.VK_A
            });
            Sanderling.KeyboardPress(VirtualKeyCode.DELETE);
            Host.Delay(500);
            Sanderling.TextEntry(itemQuantity);
            Host.Delay(1000);
            
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            var ButtonOK = Measurement?.WindowMarketAction?.FirstOrDefault()?.ButtonText?.FirstOrDefault(button=>(button?.Text).RegexMatchSuccessIgnoreCase("buy"));
            Sanderling.MouseClickLeft(ButtonOK);
            Host.Delay(5000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            CloseModalUIElementYes();
            CloseModalUIElementYes();
            CloseModalUIElementYes();
            
            //Create a new FileOrderEntry here
            double minPrice = 0.0;
            double maxPrice = CalcMinPrice(Convert.ToDouble(itemSellPrice), 0.0, defaultMargin);
            
            FileOrderEntry newFileOrder = new FileOrderEntry(itemName, "Buy Order", newPrice, minPrice, maxPrice, defaultMargin, 0, 0.0, DateTime.Now, DateTime.Now, 0, false);
            fileOrderEntries.Add(newFileOrder);

            System.IO.File.Copy(inputFileName, inputFileName + ".bak", true);
            using(FileStream fileStream2 = new FileStream(inputFileName, FileMode.Create, FileAccess.ReadWrite)) {
              using(var writer = new StreamWriter(fileStream2)) {
                List<FileOrderEntry>SortedList = fileOrderEntries.OrderByDescending(o=>o.NotFound).ThenByDescending(o=>o.OutOfPriceRange).ToList();
                foreach(FileOrderEntry fileOrderEntryToWrite in SortedList) {
                  string minPrice2 = "0.00";
                  string maxPrice2 = "0.00";
                  if (fileOrderEntryToWrite.Type.Equals("Sell Order")) {
                    minPrice2 = fileOrderEntryToWrite.LowestPrice.ToString();
                  } else {
                    maxPrice2 = fileOrderEntryToWrite.HighestPrice.ToString();
                  }
                  string output = fileOrderEntryToWrite.Name.ToString() + "," + fileOrderEntryToWrite.Type.ToString() + "," + "__" + "," + fileOrderEntryToWrite.StartPrice.ToString() + "," + "__" + "," + minPrice2 + "," + "__" + "," + maxPrice2 + "," + "__" + "," + fileOrderEntryToWrite.UpdateTime.ToString() + "," + "__" + "," + fileOrderEntryToWrite.PurchaseTime.ToString() + "," + "__" + "," + fileOrderEntryToWrite.NotFound.ToString() + "," + "__" + "," + fileOrderEntryToWrite.OutOfPriceRange.ToString();
                  writer.WriteLine(output);
                }
              }
            }
          }
        }
      }
    }
  }
} catch (Exception ex)
{
	Host.Log("Check there's not another buy window open");
	Host.Log(ex.ToString());
}

//clear the file
try {
  File.Create(orderFileName).Close();
} catch{}

public class FileOrderEntry {
  public string Name {
    get;
    set;
  }
  public string Type {
    get;
    set;
  }
  public double StartPrice {
    get;
    set;
  }
  public double LowestPrice {
    get;
    set;
  }
  public double HighestPrice {
    get;
    set;
  }
  public double Margin {
    get;
    set;
  }
  public int NumOfPriceChanges {
    get;
    set;
  }
  public double PriceChangeTotalCost {
    get;
    set;
  }
  public DateTime UpdateTime {
    get;
    set;
  }
  public DateTime PurchaseTime {
    get;
    set;
  }
  public int NotFound {
    get;
    set;
  }
  public bool OutOfPriceRange {
    get;
    set;
  }

  public FileOrderEntry(string name, string type, double startPrice, double lowestPrice, double highestPrice, double margin, int numOfPriceChanges, double priceChangeTotalCost, DateTime updateTime, DateTime purchaseTime, int notFound, bool outOfPriceRange) {
    Name = name;
    Type = type;
    StartPrice = startPrice;
    LowestPrice = lowestPrice;
    HighestPrice = highestPrice;
    Margin = margin;
    NumOfPriceChanges = numOfPriceChanges;
    PriceChangeTotalCost = priceChangeTotalCost;
    UpdateTime = updateTime;
    PurchaseTime = purchaseTime;
    NotFound = notFound;
    OutOfPriceRange = outOfPriceRange;
  }
}

double CalcMinPrice(double startPrice, double lowestPrice, double margin) {
  //If no minimum price set use margin value
  double answer = lowestPrice;
  if (lowestPrice<0.01) {
    answer = Math.Round(startPrice * (1 - (margin / 100)), 2);
  }
  return answer;
}

double CalcMaxPrice(double startPrice, double highestPrice, double margin) {
  //If no maximum price set use margin value
  double answer = highestPrice;
  if (highestPrice<0.01) {
    answer = Math.Round(startPrice * (1 + (margin / 100)), 2);
  }
  return answer;
}

bool ClickMenuEntryOnMenuRootJason(IUIElement MenuRoot, string MenuEntryRegexPattern) {
  int retryCount = 0;
  bool success = true;

  if (MenuRoot == null) {
    Host.Log("Fatal: Right click menu item is null");
  }

  Sanderling.MouseClickRight(MenuRoot);
  Host.Delay(1000);
  var Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
  var Menu = Measurement?.Menu?.FirstOrDefault();
  var MenuEntry = Menu?.EntryFirstMatchingRegexPattern(MenuEntryRegexPattern, RegexOptions.IgnoreCase);

  while (MenuEntry == null) {
    Sanderling.MouseClickRight(MenuRoot);
    Host.Delay(2000);
    Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
    Menu = Measurement?.Menu?.FirstOrDefault();
    MenuEntry = Menu?.EntryFirstMatchingRegexPattern(MenuEntryRegexPattern, RegexOptions.IgnoreCase);
    if (retryCount++>10) {
      success = false;
      break;
    }
  }

  Sanderling.MouseClickLeft(MenuEntry);
  return success;
}

void ClickMenuEntryOnMenuRoot(IUIElement MenuRoot, string MenuEntryRegexPattern) {
  Sanderling.MouseClickRight(MenuRoot);

  Host.Delay(1000);

  var Menu = Measurement?.Menu?.FirstOrDefault();

  var MenuEntry = Menu?.EntryFirstMatchingRegexPattern(MenuEntryRegexPattern, RegexOptions.IgnoreCase);

  Sanderling.MouseClickLeft(MenuEntry);
}

void EnterPrice(double price) {
  var portionInteger = (long) price;
  var cent = ((long)(price * 100)) % 100;
  Sanderling.TextEntry(portionInteger.ToString());
  Host.Delay(500);
  if (0<cent) {
    Sanderling.KeyboardPress(VirtualKeyCode.DECIMAL); // use OEM_COMMA if your locale setting has comma as decimal separator.
    Host.Delay(500);
    Sanderling.TextEntry(cent.ToString("D2"));
    Host.Delay(500);
  }
}

void CloseModalUIElementYes() {
  var ButtonClose = ModalUIElement?.ButtonText?.FirstOrDefault(button=>(button?.Text).RegexMatchSuccessIgnoreCase("yes"));

  Sanderling.MouseClickLeft(ButtonClose);
  int somethingWrong = 0;
  while(ModalUIElement != null) {
    Host.Delay(100);
    if(somethingWrong++ == 20) {
      break;
    }
  }
}

void CloseModalUIElement() {
  var ButtonClose = ModalUIElement?.ButtonText?.FirstOrDefault(button=>(button?.Text).RegexMatchSuccessIgnoreCase("close|no|ok"));

  Sanderling.MouseClickLeft(ButtonClose);
  int somethingWrong = 0;
  while(ModalUIElement != null) {
    Host.Delay(100);
    if(somethingWrong++ == 20) {
      break;
    }
  }
}

void CheckPriceColumnHeader() {
try {
  var orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.BuyerView;
  int retryCount = 0;
  while (orderSectionMarketData == null) {
    Host.Delay(500);
    if (retryCount++>20) {
      return;
    }
    Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
    orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.BuyerView;
  }
  retryCount = 0;
  while (orderSectionMarketData?.Entry?.FirstOrDefault() == null) {
    Host.Delay(500);
    if (retryCount++>20) {
      return;
    }
    Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
    orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.BuyerView;
  }            
  var firstOrder = orderSectionMarketData?.Entry?.FirstOrDefault();
  var firstPriceArray = firstOrder?.ListColumnCellLabel?.ToArray();
  string firstPriceStr = firstPriceArray[2].Value.Substring(0, firstPriceArray[2].Value.Length - 4);
  firstPriceStr = firstPriceStr.Replace(@",", "");
  double firstPriceDbl = Convert.ToDouble(firstPriceStr);
  
  bool foundBadSort = false;
  MemoryStruct.IListEntry[] buyOrders = orderSectionMarketData?.Entry?.ToArray();
  for(int i = 1; i < buyOrders.Length; i++) {
    MemoryStruct.MarketOrderEntry buyOrderInGame = (MemoryStruct.MarketOrderEntry)buyOrders[i];
    string orderText = buyOrderInGame.LabelText.FirstOrDefault().Text.ToString();
    string[] orderTextSplit = Regex.Split(orderText, @"<t>");
    string orderPrice = orderTextSplit[2];
    orderPrice = orderPrice.Replace(@"<right>", "").Replace(@",", "").Replace(@" ISK", "");
    double dblOrderPrice = Convert.ToDouble(orderPrice);
    if(firstPriceDbl < dblOrderPrice) {
      foundBadSort = true;
      break;
    }
  }
  if(foundBadSort) {
    Host.Log("Buy Price sorting incorrect.");
    Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.BuyerView.ColumnHeader[2].RegionInteraction);
  }

  orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.SellerView;
  retryCount = 0;
  while (orderSectionMarketData == null) {
    Host.Delay(500);
    if (retryCount++>20) {
      return;
    }
    Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
    orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.SellerView;
  }
  retryCount = 0;
  while (orderSectionMarketData?.Entry?.FirstOrDefault() == null) {
    Host.Delay(500);
    if (retryCount++>20) {
      return;
    }
    Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
    orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.SellerView;
  }            
  firstOrder = orderSectionMarketData?.Entry?.FirstOrDefault();
  firstPriceArray = firstOrder?.ListColumnCellLabel?.ToArray();
  firstPriceStr = firstPriceArray[2].Value.Substring(0, firstPriceArray[2].Value.Length - 4);
  firstPriceStr = firstPriceStr.Replace(@",", "");
  firstPriceDbl = Convert.ToDouble(firstPriceStr);
  
  foundBadSort = false;
  MemoryStruct.IListEntry[] sellOrders = orderSectionMarketData?.Entry?.ToArray();
  for(int i = 1; i < sellOrders.Length; i++) {
    MemoryStruct.MarketOrderEntry sellOrderInGame = (MemoryStruct.MarketOrderEntry)sellOrders[i];
    string orderText = sellOrderInGame.LabelText.FirstOrDefault().Text.ToString();
    string[] orderTextSplit = Regex.Split(orderText, @"<t>");
    string orderPrice = orderTextSplit[2];
    orderPrice = orderPrice.Replace(@"<right>", "").Replace(@",", "").Replace(@" ISK", "");
    double dblOrderPrice = Convert.ToDouble(orderPrice);
    if(firstPriceDbl > dblOrderPrice) {
      foundBadSort = true;
      break;
    }
  }
  if(foundBadSort) {
    Host.Log("Sell Price sorting incorrect.");
    Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.SellerView.ColumnHeader[2].RegionInteraction);
  }
} catch {}  
}

foundNewLoopBack:

//Click My Orders
Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.RightTabGroup?.ListTab[2]?.RegionInteraction);

//Wait for MyOrders to be populated
Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
var myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
while (myOrders == null) {
  Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
  Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.RightTabGroup?.ListTab[2]?.RegionInteraction);
  myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
  Host.Delay(5000);
}
Measurement = Sanderling?.MemoryMeasurementParsed?.Value;

for (;;) {

  //Compare the entries in my orders to the object add if not in there.
  Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
  var myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
  if (myOrders?.BuyOrderView != null && myOrders?.SellOrderView != null) {
    try {
      //If the market is open
      MemoryStruct.IListEntry[] buyOrdersInGame = myOrders?.BuyOrderView?.Entry?.ToArray();
      MemoryStruct.IListEntry[] sellOrdersInGame = myOrders?.SellOrderView?.Entry?.ToArray();
      int buyOrderCountInGame = buyOrdersInGame.Length;
      int sellOrderCountInGame = sellOrdersInGame.Length;

      //Calculate how much ISK is invested
      double totalInvested = 0.0;
      double buyInvestment = 0.0;
      double sellInvestment = 0.0;
      
      if (buyOrderCountInGame>0) {
        foreach(MemoryStruct.MarketOrderEntry buyOrderInGame in buyOrdersInGame) {
          string orderText = buyOrderInGame.LabelText.FirstOrDefault().Text.ToString();
          string[] orderTextSplit = Regex.Split(orderText, @"<t>");
          string orderQuantities = orderTextSplit[1];
          string orderPrice = orderTextSplit[2];
          string amountBought = Regex.Split(orderQuantities, @"/")[1].Replace(@"<right>", "").Replace(@",", "").Replace(@" ISK", "");
          string totalToBuy = Regex.Split(orderQuantities, @"/")[0];
          totalToBuy = totalToBuy.Replace(@"<right>", "").Replace(@",", "").Replace(@" ISK", "");
          orderPrice = orderPrice.Replace(@"<right>", "").Replace(@",", "").Replace(@" ISK", "");
          totalInvested += Convert.ToInt32(amountBought) * Convert.ToDouble(orderPrice) * 1.2;
          buyInvestment += (Convert.ToInt32(amountBought) - Convert.ToInt32(totalToBuy)) * Convert.ToDouble(orderPrice) * 1.2;
        }
      }
      if (sellOrderCountInGame>0) {
        foreach(MemoryStruct.MarketOrderEntry sellOrderInGame in sellOrdersInGame) {
          string orderText = sellOrderInGame.LabelText.FirstOrDefault().Text.ToString();
          string[] orderTextSplit = Regex.Split(orderText, @"<t>");
          string orderQuantities = orderTextSplit[1];
          string orderPrice = orderTextSplit[2];
          string totalToSell = Regex.Split(orderQuantities, @"/")[0];
          totalToSell = totalToSell.Replace(@"<right>", "").Replace(@",", "").Replace(@" ISK", "");
          orderPrice = orderPrice.Replace(@"<right>", "").Replace(@",", "").Replace(@" ISK", "");
          totalInvested += (Convert.ToInt32(totalToSell) * Convert.ToDouble(orderPrice));
          sellInvestment += (Convert.ToInt32(totalToSell) * Convert.ToDouble(orderPrice));
        }
      }
      Host.Log(@String.Format("PossProfit-{0:N}   InHanger-{1:N}   Sell-{2:N}", totalInvested, buyInvestment, sellInvestment));

      if (buyOrderCountInGame>0) {
        foreach(MemoryStruct.MarketOrderEntry buyOrderInGame in buyOrdersInGame) {
          string orderText = buyOrderInGame.LabelText.FirstOrDefault().Text.ToString();
          string[] orderTextSplit = Regex.Split(orderText, @"<t>");
          string gameOrderName = orderTextSplit[0];
          bool foundName = false;
          foreach(FileOrderEntry fileOrderEntry in fileOrderEntries) {
            // Remove () as they cause a fail
            string tempName1 = fileOrderEntry.Name.Replace(@"(", "");
            tempName1 = tempName1.Replace(@")", "");
            string tempName2 = gameOrderName.Replace(@"(", "");
            tempName2 = tempName2.Replace(@")", "");

            if (tempName1.Equals(tempName2) && fileOrderEntry.Type.Equals("Buy Order")) {
              foundName = true;
              break;
            }
          }
          if (!foundName) {
            // Add details to FileOrderEntry object
            string orderPrice = orderTextSplit[2].Substring(7, orderTextSplit[2].Length - 4 - 7);
            orderPrice = orderPrice.Replace(@",", "");
            double orderPriceDbl = Convert.ToDouble(orderPrice);
            double minPrice = CalcMinPrice(orderPriceDbl, 0.0, defaultMargin);
            double maxPrice = CalcMaxPrice(orderPriceDbl, 0.0, defaultMargin);
              
            FileOrderEntry newFileOrder = new FileOrderEntry(gameOrderName, "Buy Order", orderPriceDbl, minPrice, maxPrice, defaultMargin, 0, 0.0, DateTime.Now, DateTime.Now, 0, false);
            fileOrderEntries.Add(newFileOrder);

            //Print details for checking and pause.
            Host.Log("Name: " + newFileOrder.Name + "  Price: " + newFileOrder.StartPrice + "  Type: " + newFileOrder.Type + "  Max Price: " + newFileOrder.HighestPrice);
            //Host.Break();
            foundNew = true;
          }
        }
      }
      if (sellOrderCountInGame>0) {
        foreach(MemoryStruct.MarketOrderEntry sellOrderInGame in sellOrdersInGame) {
          string orderText = sellOrderInGame.LabelText.FirstOrDefault().Text.ToString();
          string[] orderTextSplit = Regex.Split(orderText, @"<t>");
          string gameOrderName = orderTextSplit[0];
          bool foundName = false;
          
          foreach(FileOrderEntry fileOrderEntry in fileOrderEntries) {
            if (fileOrderEntry.Name.Equals(gameOrderName) && fileOrderEntry.Type.Equals("Sell Order")) {
              foundName = true;
              break;
            }
          }
          
          if (!foundName) {
            //Get highest buying price and add defaultmargin
            if (!ClickMenuEntryOnMenuRootJason(sellOrderInGame, "View Market")) {
              Host.Log("Failed View Market Details");
            }
            int retryCount = 0;
            while (Measurement?.WindowRegionalMarket ? [0]?.SelectedItemTypeDetails?.MarketData?.BuyerView == null && Measurement?.WindowRegionalMarket ? [0]?.SelectedItemTypeDetails?.MarketData?.SellerView == null) {
              if (retryCount++>30) {
                Host.Log("Failed View Market Details");
              }
              Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
              Host.Delay(200);
            }
            var orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.BuyerView;
            while (orderSectionMarketData == null) {
              Host.Delay(500);
              Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
              orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.BuyerView;
            }
            while (orderSectionMarketData?.Entry?.FirstOrDefault() == null) {
              Host.Delay(500);
              Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
              orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.BuyerView;
            }            
            var firstOrder = orderSectionMarketData?.Entry?.FirstOrDefault();
            var firstPriceArray = firstOrder?.ListColumnCellLabel?.ToArray();
            string firstPriceStr = firstPriceArray[2].Value.Substring(0, firstPriceArray[2].Value.Length - 4);
            firstPriceStr = firstPriceStr.Replace(@",", "");
            double firstPriceDbl = Convert.ToDouble(firstPriceStr);
            firstPriceDbl = CalcMaxPrice(firstPriceDbl, 0.0, defaultMargin);
            
            //Click My Orders
            Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.RightTabGroup?.ListTab[2]?.RegionInteraction);

            //Wait for MyOrders to be populated
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
            while (myOrders == null) {
              Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
              Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.RightTabGroup?.ListTab[2]?.RegionInteraction);
              myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
              Host.Delay(1000);
            }
            
            // Add details to FileOrderEntry object
            string orderPrice = orderTextSplit[2].Substring(7, orderTextSplit[2].Length - 4 - 7);
            orderPrice = orderPrice.Replace(@",", "");
            double orderPriceDbl = Convert.ToDouble(orderPrice);
            double minPrice = CalcMinPrice(orderPriceDbl, 0.0, defaultMargin);
            double maxPrice = CalcMaxPrice(orderPriceDbl, 0.0, defaultMargin);
            
            double useMinPrice = 0.0;
            //System.Windows.Forms.DialogResult dialogResult = System.Windows.Forms.MessageBox.Show("Use auto calc value of: " + @String.Format("{0:N}", firstPriceDbl.ToString()), "Check lowest selling price", System.Windows.Forms.MessageBoxButtons.YesNo);
            //if(dialogResult == System.Windows.Forms.DialogResult.Yes)
            //{
                useMinPrice = firstPriceDbl;
            //}
            //else if (dialogResult == System.Windows.Forms.DialogResult.No)
            //{
               // useMinPrice = minPrice;
            //}              
            FileOrderEntry newFileOrder = new FileOrderEntry(gameOrderName, "Sell Order", orderPriceDbl, useMinPrice, maxPrice, defaultMargin, 0, 0.0, DateTime.Now, DateTime.Now, 0, false);
            fileOrderEntries.Add(newFileOrder);

            //Print details for checking and pause.
            Host.Log("Name: " + newFileOrder.Name + "  Price: " + newFileOrder.StartPrice + "  Type: " + newFileOrder.Type + "  Min Price: " + newFileOrder.LowestPrice);
            //Host.Break();
            foundNew = true;

            System.IO.File.Copy(inputFileName, inputFileName + ".bak", true);
            using(FileStream fileStream = new FileStream(inputFileName, FileMode.Create, FileAccess.ReadWrite)) {
              using(var writer = new StreamWriter(fileStream)) {
              List<FileOrderEntry>SortedList = fileOrderEntries.OrderByDescending(o=>o.NotFound).ThenByDescending(o=>o.OutOfPriceRange).ToList();
                foreach(FileOrderEntry fileOrderEntryToWrite in SortedList) {
                  string writeMinPrice = "0.00";
                  string writeMaxPrice = "0.00";
                  if (fileOrderEntryToWrite.Type.Equals("Sell Order")) {
                    writeMinPrice = fileOrderEntryToWrite.LowestPrice.ToString();
                  } else {
                    writeMaxPrice = fileOrderEntryToWrite.HighestPrice.ToString();
                  }
                  string output = fileOrderEntryToWrite.Name.ToString() + "," + fileOrderEntryToWrite.Type.ToString() + "," + "__" + "," + fileOrderEntryToWrite.StartPrice.ToString() + "," + "__" + "," + writeMinPrice + "," + "__" + "," + writeMaxPrice + "," + "__" + "," + fileOrderEntryToWrite.UpdateTime.ToString() + "," + "__" + "," + fileOrderEntryToWrite.PurchaseTime.ToString() + "," + "__" + "," + fileOrderEntryToWrite.NotFound.ToString() + "," + "__" + "," + fileOrderEntryToWrite.OutOfPriceRange.ToString();
                  writer.WriteLine(output);
                }
              }
            }
            
            //Something is going wrong when there are more than one new selling item.
            goto foundNewLoopBack;
          }
        }
      }
    } catch {
      /*Don't care if this fails*/
    }
  }

  Something_has_gone_wrong: //label to jump back to if something goes wrong
  int totalOrders = fileOrderEntries.Count;
  int loopCount = 1;
  
  //Sort to oldest first
  //fileOrderEntries = fileOrderEntries.OrderBy(o=>o.UpdateTime).ToList();
  fileOrderEntries = fileOrderEntries.OrderByDescending(o=>o.Type).ThenBy(o=>o.UpdateTime).ToList();
  
  foreach(FileOrderEntry fileOrderEntry in fileOrderEntries) {
    
    //If five mins and 20s has passed since last update then process again
    bool timeToCheck = false;
    DateTime a = Convert.ToDateTime(@fileOrderEntry.UpdateTime);
    DateTime b = DateTime.Now;
    if (Math.Round(b.Subtract(a).TotalSeconds, 0)>300) {
      timeToCheck = true;
    }

    if (timeToCheck && !foundNew && fileOrderEntry.NotFound < 10) {

      Host.Log("Time to check: " + fileOrderEntry.Name + " - " + fileOrderEntry.Type);

      fileOrderEntry.OutOfPriceRange = false;
      //Ensure Market Window is open
      while (null == Measurement?.WindowRegionalMarket) {
        if (null == Measurement?.WindowRegionalMarket) {
          Sanderling.MouseClickLeft(Measurement?.Neocom?.MarketButton);
        }
        Host.Delay(2000);
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
      }

      //Wait for My Orders Button to appear
      while (Measurement?.WindowRegionalMarket?.FirstOrDefault()?.RightTabGroup?.ListTab[2]?.RegionInteraction == null) {
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
        Host.Delay(500);
      }

      //Click My Orders
      Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.RightTabGroup?.ListTab[2]?.RegionInteraction);

      //Wait for MyOrders to be populated
      Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
      myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
      while (myOrders == null) {
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
        Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.RightTabGroup?.ListTab[2]?.RegionInteraction);
        myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
        Host.Delay(1000);
      }

      var orderSectionMyOrders = myOrders?.BuyOrderView;
      if (fileOrderEntry.Type.Equals("Sell Order")) {
        orderSectionMyOrders = myOrders?.SellOrderView;
      }

      //Make sure Buy/Sell section is visible
      while (orderSectionMyOrders == null) {
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
        myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
        orderSectionMyOrders = myOrders?.BuyOrderView;
        if (fileOrderEntry.Type.Equals("Sell Order")) {
          orderSectionMyOrders = myOrders?.SellOrderView;
        }
        Host.Delay(500);

        while (orderSectionMyOrders == null) {
          Host.Delay(500);
          Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
        }
      }

      //Find an order in the list that matches order read from file and View Market Details. Remove ()
      orderName = fileOrderEntry.Name.ToString();
      if (orderName.IndexOf("(")>0) orderName = fileOrderEntry.Name.ToString().Substring(0, fileOrderEntry.Name.IndexOf("("));
      var getMatchingOrder = orderSectionMyOrders?.Entry?.FirstOrDefault(MatchingOrder);

      bool foundOrder = false;
      if (getMatchingOrder != null) {
        foundOrder = true;
      } else {
        Host.Delay(1000);
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
        myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
        orderSectionMyOrders = myOrders?.BuyOrderView;
        if (fileOrderEntry.Type.Equals("Sell Order")) {
          orderSectionMyOrders = myOrders?.SellOrderView;
        }
        getMatchingOrder = orderSectionMyOrders?.Entry?.FirstOrDefault(MatchingOrder);
        if (getMatchingOrder != null) {
          foundOrder = true;
          break;
        }
        Host.Delay(1000);
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
        myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
        orderSectionMyOrders = myOrders?.BuyOrderView;
        if (fileOrderEntry.Type.Equals("Sell Order")) {
          orderSectionMyOrders = myOrders?.SellOrderView;
        }
        getMatchingOrder = orderSectionMyOrders?.Entry?.FirstOrDefault(MatchingOrder);
        if (getMatchingOrder != null) {
          foundOrder = true;
          break;
        }
        Host.Delay(1000);
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
        myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
        orderSectionMyOrders = myOrders?.BuyOrderView;
        if (fileOrderEntry.Type.Equals("Sell Order")) {
          orderSectionMyOrders = myOrders?.SellOrderView;
        }
        getMatchingOrder = orderSectionMyOrders?.Entry?.FirstOrDefault(MatchingOrder);
        if (getMatchingOrder != null) {
          foundOrder = true;
          break;
        }
        Host.Delay(1000);
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
        myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
        orderSectionMyOrders = myOrders?.BuyOrderView;
        if (fileOrderEntry.Type.Equals("Sell Order")) {
          orderSectionMyOrders = myOrders?.SellOrderView;
        }
        getMatchingOrder = orderSectionMyOrders?.Entry?.FirstOrDefault(MatchingOrder);
        if (getMatchingOrder != null) {
          foundOrder = true;
          break;
        }
        Host.Delay(1000);
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
        myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
        orderSectionMyOrders = myOrders?.BuyOrderView;
        if (fileOrderEntry.Type.Equals("Sell Order")) {
          orderSectionMyOrders = myOrders?.SellOrderView;
        }
        getMatchingOrder = orderSectionMyOrders?.Entry?.FirstOrDefault(MatchingOrder);
        if (getMatchingOrder != null) {
          foundOrder = true;
          break;
        }
        Host.Delay(1000);
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
        myOrders = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.MyOrders;
        orderSectionMyOrders = myOrders?.BuyOrderView;
        if (fileOrderEntry.Type.Equals("Sell Order")) {
          orderSectionMyOrders = myOrders?.SellOrderView;
        }
        getMatchingOrder = orderSectionMyOrders?.Entry?.FirstOrDefault(MatchingOrder);
        if (getMatchingOrder != null) {
          foundOrder = true;
          break;
        }
      }
      if (foundOrder == false) {
        Host.Log("Warning: Can't find order " + orderName);
        Console.Beep(700, 200);
        Console.Beep(700, 200);
        Console.Beep(700, 200);
        fileOrderEntry.NotFound += 1;
      } else {

        if (!ClickMenuEntryOnMenuRootJason(getMatchingOrder, "View Market")) {
          Host.Log("Failed View Market Details");
          goto Something_has_gone_wrong;
        }
        Host.Delay(2000);

        int retryCount = 0;
        while (Measurement?.WindowRegionalMarket ? [0]?.SelectedItemTypeDetails?.MarketData?.BuyerView == null && Measurement?.WindowRegionalMarket ? [0]?.SelectedItemTypeDetails?.MarketData?.SellerView == null) {
          if (retryCount++>30) {
            Host.Log("Failed View Market Details");
            goto Something_has_gone_wrong;
          }
          Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
          Host.Delay(200);
        }

        //Check that the price column headers are set correctly otherwise blue price will be off screen
        CheckPriceColumnHeader();
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
        CheckPriceColumnHeader();
        Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
	
        foreach(var button in Measurement?.WindowRegionalMarket?.FirstOrDefault()?.ButtonText) {
          if (button.Text.Equals("Export to File")) {
            Sanderling.MouseClickLeft(button.RegionInteraction);
            break;
          }
        }
        Host.Delay(1000);

        var orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.BuyerView;
        if (fileOrderEntry.Type.Equals("Sell Order")) {
          orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.SellerView;
        }
        while (orderSectionMarketData == null) {
          Host.Delay(500);
          Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
          orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.BuyerView;
          if (fileOrderEntry.Type.Equals("Sell Order")) {
            orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.SellerView;
          }
        }
        while (orderSectionMarketData?.Entry?.FirstOrDefault() == null) {
          Host.Delay(500);
          Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
          orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.BuyerView;
          if (fileOrderEntry.Type.Equals("Sell Order")) {
            orderSectionMarketData = Measurement?.WindowRegionalMarket?.FirstOrDefault()?.SelectedItemTypeDetails?.MarketData?.SellerView;
          }
        }

        //TODO: Split here for buy sell
        if (fileOrderEntry.Type.Equals("Sell Order")) {

          //Seems like a bug means that BackgroundColor is not populated on Seller unless the entry is clicked first
          Sanderling.MouseClickLeft(orderSectionMarketData?.Entry?.FirstOrDefault());
          Host.Delay(5000);

          var FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
          var FirstBlack = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlackBackground);
          //Host.Log("FirstBlue: " + string.Join(", ", FirstBlue?.ListColumnCellLabel?.Select(ColumnCellLabel=>ColumnCellLabel.Key?.Text + " : " + ColumnCellLabel.Value)?.ToArray() ?? new string[0]));
          //Host.Log("FirstBlack: " + string.Join(", ", FirstBlack?.ListColumnCellLabel?.Select(ColumnCellLabel=>ColumnCellLabel.Key?.Text + " : " + ColumnCellLabel.Value)?.ToArray() ?? new string[0]));

          bool foundBlue = false;
          if (FirstBlue != null) {
            foundBlue = true;
          } else {
            //Check multiple times to be sure
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
          }
          if (foundBlue == false) {
            Host.Log("No Blue!");
            Console.Beep(500, 1000);
            fileOrderEntry.NotFound += 1;
          }
          bool foundBlack = false;
          if (FirstBlack != null) {
            foundBlack = true;
          } else {
            //Check multiple times to be sure
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlack = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlackBackground);
            if (FirstBlack != null) {
              foundBlack = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlack = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlackBackground);
            if (FirstBlack != null) {
              foundBlack = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlack = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlackBackground);
            if (FirstBlack != null) {
              foundBlack = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlack = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlackBackground);
            if (FirstBlack != null) {
              foundBlack = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlack = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlackBackground);
            if (FirstBlack != null) {
              foundBlack = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlack = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlackBackground);
            if (FirstBlack != null) {
              foundBlack = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlack = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlackBackground);
            if (FirstBlack != null) {
              foundBlack = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlack = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlackBackground);
            if (FirstBlack != null) {
              foundBlack = true;
              break;
            }
          }

          if (foundBlack == false) {
            Host.Log("No Black!");
          }

          if (foundBlack == true && foundBlue == true) {
            var bluePriceArray = FirstBlue?.ListColumnCellLabel?.ToArray();
            string bluePriceStr = bluePriceArray[2].Value.Substring(0, bluePriceArray[2].Value.Length - 4);
            bluePriceStr = bluePriceStr.Replace(@",", "");
            double bluePriceDbl = Convert.ToDouble(bluePriceStr);

            var BlackPriceArray = FirstBlack?.ListColumnCellLabel?.ToArray();
            string BlackPriceStr = BlackPriceArray[2].Value.Substring(0, BlackPriceArray[2].Value.Length - 4);
            BlackPriceStr = BlackPriceStr.Replace(@",", "");
            double BlackPriceDbl = Convert.ToDouble(BlackPriceStr);

            double newBluePrice = 0;
            if (bluePriceDbl - BlackPriceDbl>0.005) {
              newBluePrice = BlackPriceDbl - 0.04;
              if (newBluePrice<fileOrderEntry.LowestPrice) {
                //Price gone too low
                newBluePrice = 0;
                Host.Log("Price too low on " + fileOrderEntry.Name);
                fileOrderEntry.OutOfPriceRange = true;
                Console.Beep(500, 100);
                Console.Beep(800, 100);
                Console.Beep(500, 100);
                Console.Beep(800, 100);
                Console.Beep(500, 100);
              }
            }
            try {
              if (newBluePrice>0) {
                if (!ClickMenuEntryOnMenuRootJason(FirstBlue, "Modify Order")) {
                  //Host.Log("Failed to Modify Order Sell");
                  goto Something_has_gone_wrong;
                }
                Host.Delay(2000);
                EnterPrice(newBluePrice);
                //Verify entered value
                Sanderling.InvalidateMeasurement();
                Host.Delay(2000);
                Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
                string enteredNewValueStr = Measurement?.WindowMarketAction?.FirstOrDefault()?.InputText?.FirstOrDefault()?.Text?.ToString();
                enteredNewValueStr = enteredNewValueStr.Replace(@",", "");
                double enteredNewValueDbl = Convert.ToDouble(enteredNewValueStr);

                var actionArray = Measurement?.WindowMarketAction?.FirstOrDefault()?.LabelText?.ToArray();
                string priceChangeStr = actionArray ? [12]?.Text.ToString();
                priceChangeStr = priceChangeStr.Substring(0, priceChangeStr.Length - 4);
                priceChangeStr = priceChangeStr.Replace(@",", "");
                double priceChangeDbl = Convert.ToDouble(priceChangeStr);
                priceChangeDbl = Math.Round(priceChangeDbl, 2);

                string brokerFeeStr = actionArray ? [14]?.Text.ToString();
                brokerFeeStr = brokerFeeStr.Substring(0, brokerFeeStr.Length - 4);
                brokerFeeStr = brokerFeeStr.Replace(@",", "");
                double brokerFeeDbl = Convert.ToDouble(brokerFeeStr);
                brokerFeeDbl = Math.Round(brokerFeeDbl, 2);

                //Host.Log("Entry: " + fileOrderEntry.Name + " New Price: " + newBluePrice + " Entered Price: " + enteredNewValueDbl + " Max Price: " + fileOrderEntry.HighestPrice.ToString() + " Price Change: " + (priceChangeDbl + brokerFeeDbl));

                if (Math.Abs(newBluePrice - enteredNewValueDbl)<0.011) {
                  //price is as expected so click ok
                  var ButtonOK = Measurement?.WindowMarketAction?.FirstOrDefault()?.ButtonText?.FirstOrDefault(button=>(button?.Text).RegexMatchSuccessIgnoreCase("ok"));
                  Sanderling.MouseClickLeft(ButtonOK);
                  Host.Delay(5000);
                  Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
                  CloseModalUIElementYes();
                  CloseModalUIElementYes();
                  CloseModalUIElementYes();
                  fileOrderEntry.NumOfPriceChanges = fileOrderEntry.NumOfPriceChanges + 1;
                  fileOrderEntry.UpdateTime = DateTime.Now;
                  fileOrderEntry.PriceChangeTotalCost = fileOrderEntry.PriceChangeTotalCost + priceChangeDbl + brokerFeeDbl;
                } else {
                  Host.Log("Price mismatch");                
                  var ButtonCancel = Measurement?.WindowMarketAction?.FirstOrDefault()?.ButtonText?.FirstOrDefault(button=>(button?.Text).RegexMatchSuccessIgnoreCase("cancel"));
                  Sanderling.MouseClickLeft(ButtonCancel);
                  Host.Delay(2000);
                }
              } else {
                //Host.Log("No change needed for " + orderName + " - " + fileOrderEntry.Type);
              }
            } catch {
              goto Something_has_gone_wrong;
            }
            CloseModalUIElement();
            CloseModalUIElement();
            CloseModalUIElement();
          }
        } else {
          var FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
          var FirstGreen = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsGreenBackground);
          //Host.Log("FirstBlue: " + string.Join(", ", FirstBlue?.ListColumnCellLabel?.Select(ColumnCellLabel=>ColumnCellLabel.Key?.Text + " : " + ColumnCellLabel.Value)?.ToArray() ?? new string[0]));
          //Host.Log("FirstGreen: " + string.Join(", ", FirstGreen?.ListColumnCellLabel?.Select(ColumnCellLabel=>ColumnCellLabel.Key?.Text + " : " + ColumnCellLabel.Value)?.ToArray() ?? new string[0]));

          bool foundBlue = false;
          if (FirstBlue != null) {
            foundBlue = true;
          } else {
            //Check multiple times to be sure
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstBlue = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsBlueBackground);
            if (FirstBlue != null) {
              foundBlue = true;
              break;
            }
          }
          if (foundBlue == false) {
            Host.Log("No Blue!");
            Console.Beep(500, 1000);
            fileOrderEntry.NotFound += 1;
          }

          bool foundGreen = false;
          if (FirstGreen != null) {
            foundGreen = true;
          } else {
            //Check multiple times to be sure
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstGreen = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsGreenBackground);
            if (FirstGreen != null) {
              foundGreen = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstGreen = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsGreenBackground);
            if (FirstGreen != null) {
              foundGreen = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstGreen = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsGreenBackground);
            if (FirstGreen != null) {
              foundGreen = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstGreen = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsGreenBackground);
            if (FirstGreen != null) {
              foundGreen = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstGreen = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsGreenBackground);
            if (FirstGreen != null) {
              foundGreen = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstGreen = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsGreenBackground);
            if (FirstGreen != null) {
              foundGreen = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstGreen = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsGreenBackground);
            if (FirstGreen != null) {
              foundGreen = true;
              break;
            }
            Host.Delay(1000);
            Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
            FirstGreen = orderSectionMarketData?.Entry?.FirstOrDefault(ContainsGreenBackground);
            if (FirstGreen != null) {
              foundGreen = true;
              break;
            }
          }

          if (foundGreen == false) {
            Host.Log("No Green!");
          }

          if (foundGreen == true && foundBlue == true) {
            var bluePriceArray = FirstBlue?.ListColumnCellLabel?.ToArray();
            string bluePriceStr = bluePriceArray[2].Value.Substring(0, bluePriceArray[2].Value.Length - 4);
            bluePriceStr = bluePriceStr.Replace(@",", "");
            double bluePriceDbl = Convert.ToDouble(bluePriceStr);

            var greenPriceArray = FirstGreen?.ListColumnCellLabel?.ToArray();
            string greenPriceStr = greenPriceArray[2].Value.Substring(0, greenPriceArray[2].Value.Length - 4);
            greenPriceStr = greenPriceStr.Replace(@",", "");
            double greenPriceDbl = Convert.ToDouble(greenPriceStr);
            //Host.Log("Price Diff: " + (greenPriceDbl - bluePriceDbl));
            double newBluePrice = 0;
            if (Math.Abs(greenPriceDbl - bluePriceDbl)<0.001 || greenPriceDbl - bluePriceDbl>0.005) {
              newBluePrice = greenPriceDbl + 0.04;
              if (newBluePrice>fileOrderEntry.HighestPrice) {
                //Price gone too high see if we can get next best slot
                newBluePrice = 0;
                Host.Log("Price might be too high on " + fileOrderEntry.Name);
                var entryArray = orderSectionMarketData?.Entry?.ToArray();
                fileOrderEntry.OutOfPriceRange = true;
                Console.Beep(500, 100);
                Console.Beep(800, 100);
                Console.Beep(500, 100);
                Console.Beep(800, 100);
                Console.Beep(500, 100);
                foreach(var entry in entryArray) {
                  var entryPriceArray = entry?.ListColumnCellLabel?.ToArray();
                  string entryPrice = entryPriceArray[2].Value.Substring(0, entryPriceArray[2].Value.Length - 4);
                  entryPrice = entryPrice.Replace(@",", "");
                  double entryPriceDbl = Convert.ToDouble(entryPrice);
                  if (entryPriceDbl<fileOrderEntry.HighestPrice - 1 && entryPriceDbl>bluePriceDbl) {
                    newBluePrice = entryPriceDbl + 0.02;
                    fileOrderEntry.OutOfPriceRange = false;
                    break;
                  }
                }
              }
            }
            try {
              if (newBluePrice>0) {
                if (!ClickMenuEntryOnMenuRootJason(FirstBlue, "Modify Order")) {
                  Host.Log("Failed to Modify Order Buy");
                  goto Something_has_gone_wrong;
                }
                Host.Delay(2000);
                EnterPrice(newBluePrice);
                //Verify entered value
                Sanderling.InvalidateMeasurement();
                Host.Delay(2000);
                Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
                string enteredNewValueStr = Measurement?.WindowMarketAction?.FirstOrDefault()?.InputText?.FirstOrDefault()?.Text?.ToString();
                enteredNewValueStr = enteredNewValueStr.Replace(@",", "");
                double enteredNewValueDbl = Convert.ToDouble(enteredNewValueStr);

                var actionArray = Measurement?.WindowMarketAction?.FirstOrDefault()?.LabelText?.ToArray();
                string priceChangeStr = actionArray ? [12]?.Text.ToString();
                priceChangeStr = priceChangeStr.Substring(0, priceChangeStr.Length - 4);
                priceChangeStr = priceChangeStr.Replace(@",", "");
                double priceChangeDbl = Convert.ToDouble(priceChangeStr);
                priceChangeDbl = Math.Round(priceChangeDbl, 2);

                string brokerFeeStr = actionArray ? [14]?.Text.ToString();
                brokerFeeStr = brokerFeeStr.Substring(0, brokerFeeStr.Length - 4);
                brokerFeeStr = brokerFeeStr.Replace(@",", "");
                double brokerFeeDbl = Convert.ToDouble(brokerFeeStr);
                brokerFeeDbl = Math.Round(brokerFeeDbl, 2);

                //Host.Log("Entry: " + fileOrderEntry.Name + " New Price: " + newBluePrice + " Entered Price: " + enteredNewValueDbl + " Max Price: " + fileOrderEntry.HighestPrice.ToString() + " Price Change: " + (priceChangeDbl + brokerFeeDbl));

                if (Math.Abs(newBluePrice - enteredNewValueDbl)<0.011) {
                  //price is as expected so click ok
                  var ButtonOK = Measurement?.WindowMarketAction?.FirstOrDefault()?.ButtonText?.FirstOrDefault(button=>(button?.Text).RegexMatchSuccessIgnoreCase("ok"));
                  Sanderling.MouseClickLeft(ButtonOK);
                  Host.Delay(5000);
                  Measurement = Sanderling?.MemoryMeasurementParsed?.Value;
                  CloseModalUIElementYes();
                  CloseModalUIElementYes();
                  CloseModalUIElementYes();

                  fileOrderEntry.NumOfPriceChanges = fileOrderEntry.NumOfPriceChanges + 1;
                  fileOrderEntry.UpdateTime = DateTime.Now;
                  fileOrderEntry.PriceChangeTotalCost = fileOrderEntry.PriceChangeTotalCost + priceChangeDbl + brokerFeeDbl;
                } else {
                  Host.Log("Price mismatch");                
                  var ButtonCancel = Measurement?.WindowMarketAction?.FirstOrDefault()?.ButtonText?.FirstOrDefault(button=>(button?.Text).RegexMatchSuccessIgnoreCase("cancel"));
                  Sanderling.MouseClickLeft(ButtonCancel);
                  Host.Delay(2000);
                }
              } else {
                //Host.Log("No change needed for " + orderName + " - " + fileOrderEntry.Type);
              }
            } catch {
              goto Something_has_gone_wrong;
            }
          }
        }
      }

      CloseModalUIElement();
      CloseModalUIElement();
      CloseModalUIElement();

      //Click My Orders
      Sanderling.MouseClickLeft(Measurement?.WindowRegionalMarket?.FirstOrDefault()?.RightTabGroup?.ListTab[2]?.RegionInteraction);

      //backup input file & save contents of object
      //Host.Log("Writing log.");
      System.IO.File.Copy(inputFileName, inputFileName + ".bak", true);
      using(FileStream fileStream = new FileStream(inputFileName, FileMode.Create, FileAccess.ReadWrite)) {
        using(var writer = new StreamWriter(fileStream)) {
        List<FileOrderEntry>SortedList = fileOrderEntries.OrderByDescending(o=>o.NotFound).ThenByDescending(o=>o.OutOfPriceRange).ToList();
          foreach(FileOrderEntry fileOrderEntryToWrite in SortedList) {
            string minPrice = "0.00";
            string maxPrice = "0.00";
            if (fileOrderEntryToWrite.Type.Equals("Sell Order")) {
              minPrice = fileOrderEntryToWrite.LowestPrice.ToString();
            } else {
              maxPrice = fileOrderEntryToWrite.HighestPrice.ToString();
            }
            string output = fileOrderEntryToWrite.Name.ToString() + "," + fileOrderEntryToWrite.Type.ToString() + "," + "__" + "," + fileOrderEntryToWrite.StartPrice.ToString() + "," + "__" + "," + minPrice + "," + "__" + "," + maxPrice + "," + "__" + "," + fileOrderEntryToWrite.UpdateTime.ToString() + "," + "__" + "," + fileOrderEntryToWrite.PurchaseTime.ToString() + "," + "__" + "," + fileOrderEntryToWrite.NotFound.ToString() + "," + "__" + "," + fileOrderEntryToWrite.OutOfPriceRange.ToString();
            writer.WriteLine(output);
          }
        }
      }
    }
    Host.Log("Done: " + loopCount++.ToString() + " of " + totalOrders);
  }

  if (foundNew == true) {
    foundNew = false;
    //backup input file & save contents of object
    //Host.Log("Writing log.");
    System.IO.File.Copy(inputFileName, inputFileName + ".bak", true);
    using(FileStream fileStream = new FileStream(inputFileName, FileMode.Create, FileAccess.ReadWrite)) {
      using(var writer = new StreamWriter(fileStream)) {
        List<FileOrderEntry>SortedList = fileOrderEntries.OrderByDescending(o=>o.NotFound).ThenByDescending(o=>o.OutOfPriceRange).ToList();
        foreach(FileOrderEntry fileOrderEntryToWrite in SortedList) {
          string minPrice = "0.00";
          string maxPrice = "0.00";
          if (fileOrderEntryToWrite.Type.Equals("Sell Order")) {
            minPrice = fileOrderEntryToWrite.LowestPrice.ToString();
          } else {
            maxPrice = fileOrderEntryToWrite.HighestPrice.ToString();
          }
          string output = fileOrderEntryToWrite.Name.ToString() + "," + fileOrderEntryToWrite.Type.ToString() + "," + "__" + "," + fileOrderEntryToWrite.StartPrice.ToString() + "," + "__" + "," + minPrice + "," + "__" + "," + maxPrice + "," + "__" + "," + fileOrderEntryToWrite.UpdateTime.ToString() + "," + "__" + "," + fileOrderEntryToWrite.PurchaseTime.ToString() + "," + "__" + "," + fileOrderEntryToWrite.NotFound.ToString() + "," + "__" + "," + fileOrderEntryToWrite.OutOfPriceRange.ToString();
          writer.WriteLine(output);
        }
      }
    }
  }

  //Check every few seconds for item to be checked
  Console.Beep(700, 200);
  int delay = rnd.Next(12423, 15624);
  Host.Delay(delay);
}
