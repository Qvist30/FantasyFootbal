/**
 * The Week class represents results for a team in a given week.
 * 
 * @author Tom Kennedy
 */
public class Week {

	/**
	 * There can be only three outcomes in a given week. You can win, lose or
	 * tie the game.
	 */
	public enum Outcome {
		WIN, LOSS, TIE
	};

	/** The week number of the season. */
	private int weekNumber;

	/** The score that the team had this week. */
	private int score;

	/** The outcome of this week's game. */
	private Outcome result;

	/** How the team would have fared against all teams this week. */
	private float totalRecord;

	/**
	 * Gets the number of week that this was.
	 * 
	 * @return The number of week that this was.
	 */
	public int getWeekNumber() {
		return weekNumber;
	}

	/**
	 * Sets the week number.
	 * 
	 * @param weekNumber
	 *            The week number.
	 */
	public void setWeekNumber(int weekNumber) {
		this.weekNumber = weekNumber;
	}

	/**
	 * Gets the team's score for the week.
	 * 
	 * @return The team's score for the week.
	 */
	public int getScore() {
		return score;
	}

	/**
	 * Sets the team's score for the week.
	 * 
	 * @param score
	 *            The team's score for the week.
	 */
	public void setScore(int score) {
		this.score = score;
	}

	/**
	 * Gets the outcome of the game: Win, Loss, or Tie.
	 * 
	 * @return The outcome of the game.
	 */
	public Outcome getResult() {
		return result;
	}

	/**
	 * Sets the outcome of the game.
	 * 
	 * @param result
	 *            The outcome of the game.
	 */
	public void setResult(Outcome result) {
		this.result = result;
	}

	/**
	 * Gets the team's record if it played every team this week.
	 * 
	 * @return The total record.
	 */
	public float getTotalRecord() {
		return totalRecord;
	}

	/**
	 * Sets the team's record if it played every team this week.
	 * 
	 * @param totalRecord
	 *            The total record.
	 */
	public void setTotalRecord(float totalRecord) {
		this.totalRecord = totalRecord;
	}

	/**
	 * Gets the weighted points, which is the week number times the number of
	 * points. The theory behind this is the more recent weeks should count
	 * heavier.
	 * 
	 * @return The weighted points for the week.
	 */
	public int getWeightedPoints() {
		return score * weekNumber;
	}
}
