public class Utils {

    public static int toNumber(String name) {
        int number = 0;
        for (int i = 0; i < name.length(); i++) {
            number = number * 26 + (name.charAt(i) - ('A' - 1));
        }
        return number;
    }

    public static String toName(int number) {
        StringBuilder sb = new StringBuilder();
        while (number-- > 0) {
            sb.append((char) ('A' + (number % 26)));
            number /= 26;
        }
        return sb.reverse().toString();
    }

    public static String getLetter(String value) {
        StringBuilder strAZ = new StringBuilder();
        for (int i = 0; i < value.length(); i++) {
            Character charValue = value.charAt(i);
            if (Character.isLetter(charValue)) {
                strAZ.append(charValue);
            }
        }
        return strAZ.toString();
    }

    public static int getDigit(String value) {
        StringBuilder str09 = new StringBuilder();
        for (int i = 0; i < value.length(); i++) {
            Character charValue = value.charAt(i);
            if (Character.isDigit(charValue)) {
                str09.append(charValue);
            }
        }
        return Integer.parseInt(str09.toString());
    }
}
