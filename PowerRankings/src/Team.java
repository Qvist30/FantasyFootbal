import java.util.ArrayList;
import java.util.List;

/**
 * A Team object represents a fantasy football team, and contains all of the
 * information pertaining to its week-to-week results.
 * 
 * @author Tom Kennedy
 */
public class Team implements Comparable<Team> {

	/** The name of the player who owns the team. */
	private String name;

	/** All of the weeks that this team has competed in. */
	private List<Week> weeks = new ArrayList<Week>();

	/**
	 * Constructor for Team.
	 * 
	 * @param name
	 *            The name of the player that owns the team.
	 */
	public Team(String name) {
		this.name = name;
	}

	/**
	 * Adds a week to the season that this team has competed in.
	 * 
	 * @param week
	 *            The week that this team has played.
	 */
	public void addWeek(Week week) {
		weeks.add(week);
	}

	/**
	 * Prints a team's record in format of Wins-Losses. If the team has any ties
	 * the format will be Wins-Losses-Ties.
	 * 
	 * @return A nicely formatted record.
	 */
	public String prettyPrintRecord() {
		int wins = 0;
		int losses = 0;
		int ties = 0;
		for (Week week : weeks) {
			switch (week.getResult()) {
			case WIN:
				wins++;
				break;
			case LOSS:
				losses++;
				break;
			case TIE:
				ties++;
				break;
			}
		}
		StringBuilder builder = new StringBuilder();
		builder.append(wins);
		builder.append('-');
		builder.append(losses);
		if (ties != 0) {
			builder.append('-');
			builder.append(ties);
		}
		return builder.toString();
	}

	/**
	 * A programmitc representation of the team's record. Ties count as half a
	 * win.
	 * 
	 * @return The record of the team.
	 */
	public float getRecord() {
		float wins = 0;
		for (Week week : weeks) {
			if (week.getResult() == Week.Outcome.WIN) {
				wins++;
			} else if (week.getResult() == Week.Outcome.TIE) {
				wins += 0.5;
			}
		}
		return wins;
	}

	/**
	 * Gets the name of the player that owns this team.
	 * 
	 * @return The name of the player that owns this team.
	 */
	public String getName() {
		return name;
	}

	/**
	 * Gets the record vs. all teams. This represents how the team would be
	 * doing if it played every team every single week.
	 * 
	 * @return The record vs. all teams in the league.
	 */
	public float getRecordVsAll() {
		float recordVsAll = 0;
		for (Week week : weeks) {
			recordVsAll += week.getTotalRecord();
		}
		return recordVsAll;
	}

	/**
	 * Gets the sum of all of the weighted points that this team has accrued throughout the season.
	 * 
	 * @return The sum of all of this teams weighted points.
	 */
	public int getWeightedPoints() {
		int totalWeightedPoints = 0;
		for (Week week : weeks) {
			totalWeightedPoints += week.getWeightedPoints();
		}
		return totalWeightedPoints;
	}

	/**
	 * Compares a team against this team. If a team wins two categories it is
	 * the better team.
	 * 
	 * @see java.lang.Comparable#compareTo(java.lang.Object)
	 */
	public int compareTo(Team team2) {
		final int BEFORE = -1;
		final int EQUAL = 0;
		final int AFTER = 1;
		int comparison;
		if (this == team2) {
			comparison = EQUAL;
		} else {
			if (this.getRecord() > team2.getRecord()
					&& (this.getRecordVsAll() > team2.getRecordVsAll() || this
							.getWeightedPoints() > team2.getWeightedPoints())) {
				comparison = BEFORE;
			} else if (this.getRecordVsAll() > team2.getRecordVsAll()
					&& this.getWeightedPoints() > team2.getWeightedPoints()) {
				comparison = BEFORE;
			} else {
				comparison = AFTER;
			}
		}
		return comparison;
	}
}
