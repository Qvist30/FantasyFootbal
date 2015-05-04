import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.StringTokenizer;

/**
 * This class takes an already created .csv file with a set format of <Name, points, Vs. Ranking, Win, points, Vs. Ranking, Win...>.
 * and prints out the record of each team, as well as how the rate overall in the power rankings.
 * 
 * @author Tom Kennedy
 */
public class FantasyFootballSample {
	
	/** String constant to recognize Wins. */
	private static final String WIN = "1";
	
	/** String constant to recognize Ties. */
	private static final String TIE = "0.5";

	/** All of the teams that are found in the file. */
	private static List<Team> teams = new ArrayList<Team>();

	/**
	 * Takes a .csv file and prints out the record of each team, as well as the total power ranking for each team.
	 * 
	 * @param args args[0] must be the .csv file.  Any additional arguments will be ignored.
	 */
	public static void main(String[] args) throws Exception {
		// power rankings csv file.
		File powerRankingsFile;
		if(args.length == 0 || args[0] == null) {
			// Default file location.
			powerRankingsFile = new File("etc/power_rankings.csv");
		} else {
			powerRankingsFile = new File(args[0]);
		}
		BufferedReader br = new BufferedReader(new FileReader(powerRankingsFile));
		
		String string;
		int weekNumber = 1;
		while((string = br.readLine()) != null && !string.trim().equals("")) {
			StringTokenizer tokenizer = new StringTokenizer(string, ",");
			// Team names are always the first token.
			Team team = new Team(tokenizer.nextToken());
			teams.add(team);
			while(tokenizer.hasMoreTokens()) {
				Week week = new Week();
				week.setWeekNumber(weekNumber);

				// First token will be score
				week.setScore(Integer.valueOf(tokenizer.nextToken()));
				
				// Second token is the record vs. all teams.
				week.setTotalRecord(Float.valueOf(tokenizer.nextToken()));
				
				// Third token is Win, Lose, or Draw.
				String outcome = tokenizer.nextToken();
				if(WIN.equals(outcome)) {
					week.setResult(Week.Outcome.WIN);
				} else if(TIE.equals(outcome)) {
					week.setResult(Week.Outcome.TIE);
				} else {
					week.setResult(Week.Outcome.LOSS);
				}
				
				weekNumber++;
				team.addWeek(week);
			}
			weekNumber = 0;
			System.out.println(team.getName() + " had a record of " + team.prettyPrintRecord() + ".");
		}
		determinePowerRankings(teams);
	}

	/**
	 * Determines and prints the power rankings by sorting the list of teams.
	 * 
	 * @param teams The teams that have been found in the file.
	 */
	private static void determinePowerRankings(List<Team> teams) {
		Collections.sort(teams);
		for(int i=0; i<teams.size(); i++) {
			System.out.println(teams.get(i).getName() + " is ranked #" + (i+1) + " overall." );
		}
	}
}
