import java.util.Scanner;

public class ConsoleQuestion {

    public  static boolean getAnswer(String question) {
        System.out.println(TextColor.BLUE
                + question
                + TextColor.RESET + "[ "
                + TextColor.GREEN + "Y"
                + TextColor.RESET + " / "
                + TextColor.RED + "N"
                + TextColor.RESET + " ]");
        return "y".equals(new Scanner(System.in).next().toLowerCase().replace(" ", ""));
    }
}
