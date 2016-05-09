/*PLEASE DO NOT EDIT THIS CODE*/
/*This code was generated using the UMPLE 1.22.0.5146 modeling language!*/

package ca.mcgilleus.budgetbuilder.model;
import java.util.*;

import org.apache.poi.ss.usermodel.CellStyle;

// line 9 "../../../../Model.ump"
// line 30 "../../../../Model.ump"
public class Portfolio
{

  //------------------------
  // MEMBER VARIABLES
  //------------------------

  //Portfolio Attributes
  private String name;
  private CellStyle portfolioLabelStyle;

  //Portfolio Associations
  private List<CommitteeBudget> committeeBudgets;
  private EUSBudget eUSBudget;

  //------------------------
  // CONSTRUCTOR
  //------------------------

  public Portfolio(String aName, EUSBudget aEUSBudget)
  {
    name = aName;
    committeeBudgets = new ArrayList<CommitteeBudget>();
    boolean didAddEUSBudget = setEUSBudget(aEUSBudget);
    if (!didAddEUSBudget)
    {
      throw new RuntimeException("Unable to create portfolio due to eUSBudget");
    }
  }

  //------------------------
  // INTERFACE
  //------------------------

  public boolean setName(String aName)
  {
    boolean wasSet = false;
    name = aName;
    wasSet = true;
    return wasSet;
  }

  public String getName()
  {
    return name;
  }

  public CommitteeBudget getCommitteeBudget(int index)
  {
    CommitteeBudget aCommitteeBudget = committeeBudgets.get(index);
    return aCommitteeBudget;
  }

  public List<CommitteeBudget> getCommitteeBudgets()
  {
    List<CommitteeBudget> newCommitteeBudgets = Collections.unmodifiableList(committeeBudgets);
    return newCommitteeBudgets;
  }

  public int numberOfCommitteeBudgets()
  {
    int number = committeeBudgets.size();
    return number;
  }

  public boolean hasCommitteeBudgets()
  {
    boolean has = committeeBudgets.size() > 0;
    return has;
  }

  public int indexOfCommitteeBudget(CommitteeBudget aCommitteeBudget)
  {
    int index = committeeBudgets.indexOf(aCommitteeBudget);
    return index;
  }

  public EUSBudget getEUSBudget()
  {
    return eUSBudget;
  }

  public static int minimumNumberOfCommitteeBudgets()
  {
    return 0;
  }

  public CommitteeBudget addCommitteeBudget(String aName, String amtRequestedRef)
  {
    return new CommitteeBudget(aName, amtRequestedRef, this);
  }

  public boolean addCommitteeBudget(CommitteeBudget aCommitteeBudget)
  {
    boolean wasAdded = false;
    if (committeeBudgets.contains(aCommitteeBudget)) { return false; }
    Portfolio existingPortfolio = aCommitteeBudget.getPortfolio();
    boolean isNewPortfolio = existingPortfolio != null && !this.equals(existingPortfolio);
    if (isNewPortfolio)
    {
      aCommitteeBudget.setPortfolio(this);
    }
    else
    {
      committeeBudgets.add(aCommitteeBudget);
    }
    wasAdded = true;
    return wasAdded;
  }

  public boolean removeCommitteeBudget(CommitteeBudget aCommitteeBudget)
  {
    boolean wasRemoved = false;
    //Unable to remove aCommitteeBudget, as it must always have a portfolio
    if (!this.equals(aCommitteeBudget.getPortfolio()))
    {
      committeeBudgets.remove(aCommitteeBudget);
      wasRemoved = true;
    }
    return wasRemoved;
  }

  public boolean addCommitteeBudgetAt(CommitteeBudget aCommitteeBudget, int index)
  {  
    boolean wasAdded = false;
    if(addCommitteeBudget(aCommitteeBudget))
    {
      if(index < 0 ) { index = 0; }
      if(index > numberOfCommitteeBudgets()) { index = numberOfCommitteeBudgets() - 1; }
      committeeBudgets.remove(aCommitteeBudget);
      committeeBudgets.add(index, aCommitteeBudget);
      wasAdded = true;
    }
    return wasAdded;
  }

  public boolean addOrMoveCommitteeBudgetAt(CommitteeBudget aCommitteeBudget, int index)
  {
    boolean wasAdded = false;
    if(committeeBudgets.contains(aCommitteeBudget))
    {
      if(index < 0 ) { index = 0; }
      if(index > numberOfCommitteeBudgets()) { index = numberOfCommitteeBudgets() - 1; }
      committeeBudgets.remove(aCommitteeBudget);
      committeeBudgets.add(index, aCommitteeBudget);
      wasAdded = true;
    } 
    else 
    {
      wasAdded = addCommitteeBudgetAt(aCommitteeBudget, index);
    }
    return wasAdded;
  }

  public boolean setEUSBudget(EUSBudget aEUSBudget)
  {
    boolean wasSet = false;
    if (aEUSBudget == null)
    {
      return wasSet;
    }

    EUSBudget existingEUSBudget = eUSBudget;
    eUSBudget = aEUSBudget;
    if (existingEUSBudget != null && !existingEUSBudget.equals(aEUSBudget))
    {
      existingEUSBudget.removePortfolio(this);
    }
    eUSBudget.addPortfolio(this);
    wasSet = true;
    return wasSet;
  }

  public void delete()
  {
    for(int i=committeeBudgets.size(); i > 0; i--)
    {
      CommitteeBudget aCommitteeBudget = committeeBudgets.get(i - 1);
      aCommitteeBudget.delete();
    }
    EUSBudget placeholderEUSBudget = eUSBudget;
    this.eUSBudget = null;
    placeholderEUSBudget.removePortfolio(this);
  }


  public String toString()
  {
	  String outputString = "";
    return super.toString() + "["+
            "name" + ":" + getName()+ "]" + System.getProperties().getProperty("line.separator") +
            "  " + "eUSBudget = "+(getEUSBudget()!=null?Integer.toHexString(System.identityHashCode(getEUSBudget())):"null")
     + outputString;
  }

public CellStyle getPortfolioLabelStyle() {
	return portfolioLabelStyle;
}

public void setPortfolioLabelStyle(CellStyle portfolioLabelStyle) {
	this.portfolioLabelStyle = portfolioLabelStyle;
}
}