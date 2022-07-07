= Table.AddColumn(
    #"Body Expandido",
    "Related Number",
    each List.Distinct(
        List.Select(
            List.Transform(
                Splitter.SplitTextByCharacterTransition(
                    {
                        Character.FromNumber(0)..Character.FromNumber(47),
                        Character.FromNumber(58)..Character.FromNumber(1000)
                    },
                    {"0".."9"}
                )([Subject] & [Body.TextBody]),
                each Text.Select(_, {"0".."9"})
            ),
            each Text.Length(_) > 8 and Text.Length(_) < 13
        )
    )
)