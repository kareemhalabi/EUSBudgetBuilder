package ca.mcgilleus.budgetbuilder.model;

/*PLEASE DO NOT EDIT THIS CODE*/

/*This code was generated using the UMPLE 1.24.0-6d4c262 modeling language!*/

// line 13 "model.ump"
// line 37 "model.ump"
public class CommitteeBudget {

	// ------------------------
	// MEMBER VARIABLES
	// ------------------------

	// CommitteeBudget Attributes
	private String name;
	private String amtRequestedRef;

	// CommitteeBudget Associations
	private Portfolio portfolio;

	// ------------------------
	// CONSTRUCTOR
	// ------------------------

	public CommitteeBudget(String aName, String aAmtRequestedRef, Portfolio aPortfolio) {
		name = aName;
		amtRequestedRef = aAmtRequestedRef;
		boolean didAddPortfolio = setPortfolio(aPortfolio);
		if (!didAddPortfolio) {
			throw new RuntimeException("Unable to create committeeBudget due to portfolio");
		}
	}

	// ------------------------
	// INTERFACE
	// ------------------------

	public boolean setName(String aName) {
		boolean wasSet = false;
		name = aName;
		wasSet = true;
		return wasSet;
	}

	public boolean setAmtRequestedRef(String aAmtRequestedRef) {
		boolean wasSet = false;
		amtRequestedRef = aAmtRequestedRef;
		wasSet = true;
		return wasSet;
	}

	public String getName() {
		return name;
	}

	public String getAmtRequestedRef() {
		return amtRequestedRef;
	}

	public Portfolio getPortfolio() {
		return portfolio;
	}

	public boolean setPortfolio(Portfolio aPortfolio) {
		boolean wasSet = false;
		if (aPortfolio == null) {
			return wasSet;
		}

		Portfolio existingPortfolio = portfolio;
		portfolio = aPortfolio;
		if (existingPortfolio != null && !existingPortfolio.equals(aPortfolio)) {
			existingPortfolio.removeCommitteeBudget(this);
		}
		portfolio.addCommitteeBudget(this);
		wasSet = true;
		return wasSet;
	}

	public void delete() {
		Portfolio placeholderPortfolio = portfolio;
		this.portfolio = null;
		placeholderPortfolio.removeCommitteeBudget(this);
	}

	public String toString() {
		String outputString = "";
		return super.toString() + "[" + "name" + ":" + getName() + "," + "amtRequestedRef" + ":" + getAmtRequestedRef()
				+ "]" + System.getProperties().getProperty("line.separator") + "  " + "portfolio = "
				+ (getPortfolio() != null ? Integer.toHexString(System.identityHashCode(getPortfolio())) : "null")
				+ outputString;
	}
}