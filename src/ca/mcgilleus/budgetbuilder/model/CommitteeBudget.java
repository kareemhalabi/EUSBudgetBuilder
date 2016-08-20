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
	private String revRef;
	private String expRef;

	private double previousAmt;

	// CommitteeBudget Associations
	private Portfolio portfolio;

	// ------------------------
	// CONSTRUCTOR
	// ------------------------

	public CommitteeBudget(String aName, String aRevRef, String aExpRef, Portfolio aPortfolio) {
		name = aName;
		revRef = aRevRef;
		expRef = aExpRef;
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

	public boolean setRevRef(String aRevRef) {
		boolean wasSet = false;
		revRef = aRevRef;
		wasSet = true;
		return wasSet;
	}

	public boolean setExpRef(String aExpRef) {
		boolean wasSet = false;
		expRef = aExpRef;
		wasSet = true;
		return wasSet;
	}

	public boolean setPreviousAmt(double aPreviousAmt) {
		boolean wasSet = false;
		previousAmt = aPreviousAmt;
		wasSet = true;
		return wasSet;
	}

	public String getName() {
		return name;
	}

	public String getRevRef() {
		return revRef;
	}

	public String getExpRef() {
		return expRef;
	}

	public double getPreviousAmt() {
		return previousAmt;
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
		return super.toString() + "[" + "name" + ":" + getName() + "," + "revRef" + ":" + getRevRef()
				+ "expRef" + ":" + getExpRef() + "]" + System.getProperties().getProperty("line.separator") + "  " + "portfolio = "
				+ (getPortfolio() != null ? Integer.toHexString(System.identityHashCode(getPortfolio())) : "null")
				+ outputString;
	}

	@Override
	public boolean equals(Object other) {
		if(other instanceof CommitteeBudget) {

			CommitteeBudget otherCommittee = (CommitteeBudget) other;
			return this.getName().trim().toUpperCase().equals(
					otherCommittee.getName().trim().toUpperCase())
				&& 	this.getPortfolio().equals(otherCommittee.getPortfolio());

		} else if(other instanceof String) {
			String otherName = (String) other;
			return this.getName().trim().toUpperCase().equals(
					otherName.trim().toUpperCase()
			);
		}
		return false;
	}
}