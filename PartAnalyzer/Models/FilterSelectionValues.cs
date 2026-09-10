namespace PartAnalyzer.Models;

public static class FilterSelectionValues
{
    // Excel cell text cannot contain this control character, so it cannot collide with a source Category value.
    public const string BlankCategory = "\u0001PartAnalyzerBlankCategory";
}
