/*PLEASE DO NOT EDIT THIS CODE*/
/*This code was generated using the UMPLE 1.22.0.5146 modeling language!*/

package ca.mcgilleus.budgetbuilder.model;
import java.util.Date;
import java.util.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// line 3 "../../../../Model.ump"
// line 18 "../../../../Model.ump"
// line 23 "../../../../Model.ump"
public class EUSBudget
{

  //------------------------
  // MEMBER VARIABLES
  //------------------------

  //EUSBudget Attributes
  private Date year;
  private XSSFWorkbook budget;

  //EUSBudget Associations
  private List<Portfolio> portfolios;
  private EUSBudget previousYear;

  //------------------------
  // CONSTRUCTOR
  //------------------------

  public EUSBudget(Date aYear)
  {
    year = aYear;
    portfolios = new ArrayList<Portfolio>();
    budget = new XSSFWorkbook();
  }

  //------------------------
  // INTERFACE
  //------------------------

  public boolean setYear(Date aYear)
  {
    boolean wasSet = false;
    year = aYear;
    wasSet = true;
    return wasSet;
  }

  public Date getYear()
  {
    return year;
  }

  public XSSFWorkbook getBudget() {
	return budget;
}

public Portfolio getPortfolio(int index)
  {
    Portfolio aPortfolio = portfolios.get(index);
    return aPortfolio;
  }

  public List<Portfolio> getPortfolios()
  {
    List<Portfolio> newPortfolios = Collections.unmodifiableList(portfolios);
    return newPortfolios;
  }

  public int numberOfPortfolios()
  {
    int number = portfolios.size();
    return number;
  }

  public boolean hasPortfolios()
  {
    boolean has = portfolios.size() > 0;
    return has;
  }

  public int indexOfPortfolio(Portfolio aPortfolio)
  {
    int index = portfolios.indexOf(aPortfolio);
    return index;
  }

  public EUSBudget getPreviousYear()
  {
    return previousYear;
  }

  public boolean hasPreviousYear()
  {
    boolean has = previousYear != null;
    return has;
  }

  public static int minimumNumberOfPortfolios()
  {
    return 0;
  }

  public Portfolio addPortfolio(String aName)
  {
    return new Portfolio(aName, this);
  }

  public boolean addPortfolio(Portfolio aPortfolio)
  {
    boolean wasAdded = false;
    if (portfolios.contains(aPortfolio)) { return false; }
    EUSBudget existingEUSBudget = aPortfolio.getEUSBudget();
    boolean isNewEUSBudget = existingEUSBudget != null && !this.equals(existingEUSBudget);
    if (isNewEUSBudget)
    {
      aPortfolio.setEUSBudget(this);
    }
    else
    {
      portfolios.add(aPortfolio);
    }
    wasAdded = true;
    return wasAdded;
  }

  public boolean removePortfolio(Portfolio aPortfolio)
  {
    boolean wasRemoved = false;
    //Unable to remove aPortfolio, as it must always have a eUSBudget
    if (!this.equals(aPortfolio.getEUSBudget()))
    {
      portfolios.remove(aPortfolio);
      wasRemoved = true;
    }
    return wasRemoved;
  }

  public boolean addPortfolioAt(Portfolio aPortfolio, int index)
  {  
    boolean wasAdded = false;
    if(addPortfolio(aPortfolio))
    {
      if(index < 0 ) { index = 0; }
      if(index > numberOfPortfolios()) { index = numberOfPortfolios() - 1; }
      portfolios.remove(aPortfolio);
      portfolios.add(index, aPortfolio);
      wasAdded = true;
    }
    return wasAdded;
  }

  public boolean addOrMovePortfolioAt(Portfolio aPortfolio, int index)
  {
    boolean wasAdded = false;
    if(portfolios.contains(aPortfolio))
    {
      if(index < 0 ) { index = 0; }
      if(index > numberOfPortfolios()) { index = numberOfPortfolios() - 1; }
      portfolios.remove(aPortfolio);
      portfolios.add(index, aPortfolio);
      wasAdded = true;
    } 
    else 
    {
      wasAdded = addPortfolioAt(aPortfolio, index);
    }
    return wasAdded;
  }

  public boolean setPreviousYear(EUSBudget aNewPreviousYear)
  {
    boolean wasSet = false;
    previousYear = aNewPreviousYear;
    wasSet = true;
    return wasSet;
  }

  public void delete()
  {
    for(int i=portfolios.size(); i > 0; i--)
    {
      Portfolio aPortfolio = portfolios.get(i - 1);
      aPortfolio.delete();
    }
    previousYear = null;
  }


  public String toString()
  {
	  String outputString = "";
    return super.toString() + "["+ "]" + System.getProperties().getProperty("line.separator") +
            "  " + "year" + "=" + (getYear() != null ? !getYear().equals(this)  ? getYear().toString().replaceAll("  ","    ") : "this" : "null")
     + outputString;
  }
}