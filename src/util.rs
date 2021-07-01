/// used for convenient access to the "A" side of "A1" notation
///
/// This is hard-coded because it'll compile to binary constant and be really nice and fast.
///
/// Take for example A1:C3, a range that spreads over 3 rows and (_1:_3) and 3 columns (A_:C_)
/// If we wanted to specify this range and we have a `Vec<Vec<u32>>` specifying the values
/// that should be placed in this range, we could make that range like so:
///
/// ```rust
/// const ASCII_UPPER: [char; 26] = [
///     'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
///     'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
/// ];
/// let data = vec![vec![1, 2, 3,], vec![4, 5, 6], vec![7, 8, 9]];
///
/// let start_column = ASCII_UPPER[0];
/// let end_column = ASCII_UPPER[data[0].len()];
/// let start_row = 0;
/// let end_row = data.len();
/// let range = format!("{}{}:{}{}", start_column, start_row, end_column, end_row);
/// println!("{}", range);  // -> A1:C3
/// ```
const ASCII_UPPER: [char; 26] = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S',
    'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
];

/// helper function to get column notation ("A", "CF") from a zero-indexed number
///
/// For instance, the first column in a google sheets page is "A".
///
/// ```ignore
/// let first_column_notation = get_column_notation(0);
/// println!("{}", &first_column_notation); // -> "A"
///
/// let further_column = get_column_notation(27);
/// println!("{}", &further_column);  // -> "AB";
///```
fn get_column_notation(column: usize) -> String {
    // A - Z
    if column < 26 {
        return format!("{}", ASCII_UPPER[column]);
    }
    // AA - ZZ
    else if column < 702 {
        let one = ASCII_UPPER[column / 26 - 1];
        let two = ASCII_UPPER[column % 26];
        format!("{}{}", one, two)
    }
    // AAA - ZZZ
    else if column <= 18277 {
        // honestly don't understand why this works
        let first = if column / 26 / 26 >= 26 {
            26 - (column / 26 / 26) % 26
        } else {
            column / 26 / 26 - 1
        };
        let one = ASCII_UPPER[first];
        let two = ASCII_UPPER[(column / 26 - 1) % 26];
        let three = ASCII_UPPER[column % 26];
        format!("{}{}{}", one, two, three)
    }
    // Number not supported — Over Column ZZZ
    else {
        panic!("Number not supported. A number higher than 18277 will go beyond Column ZZZ.");
    }
}

/// This is a helper function to retrieve valid A1 notation given the starting and ending index for
/// columns and rows in a zero-index fashion. This is used in zero-index fashion to make it easier to work
/// with arrays `Vec` of data!
///
/// Please refer to [Google Sheets Docs: A1 Notation] for more information on A1 Notation
///
/// # Examples
///
/// ```rust
/// use googlesheets::util::get_a1_notation;
///
/// let top_left_nine_cells = get_a1_notation(Some(0), Some(0), Some(2), Some(2));
/// println!("{}", top_left_nine_cells);  // -> "A1:C3"
///
/// let rows_five_through_nine = get_a1_notation(None, Some(4), None, Some(8));
/// println!("{}", rows_five_through_nine); // -> "5:9"
///
/// let rows_five_through_nine_third_column_on = get_a1_notation(None, Some(4), Some(2), Some(8));
/// println!("{}", rows_five_through_nine_third_column_on); // -> "5:C9"
/// ```
///
/// [Google Sheets Docs: A1 Notation]: https://developers.google.com/sheets/api/guides/concepts#expandable-1
pub fn get_a1_notation(
    start_column: Option<usize>,
    start_row: Option<usize>,
    end_column: Option<usize>,
    end_row: Option<usize>,
) -> String {
    match (start_column, start_row, end_column, end_row) {
        // "A5:A" refers to all the cells in the first column, from row 5 onward
        (Some(sc), Some(r), Some(ec), None) |
        // "A:A5" is not technically valid, but defaults to "A5:A"
        (Some(sc), None, Some(ec), Some(r)) => {
            format!("{}{}:{}", get_column_notation(sc), r+1, get_column_notation(ec))
        },
        // "A1:B2" refers to the first two cells in the top two rows
        (Some(sc), Some(sr), Some(ec), Some(er)) => {
            format!("{}{}:{}{}", get_column_notation(sc), sr+1, get_column_notation(ec), er+1)
        },
        // "A:B" refers to all the cells in the first two columns
        (Some(sc), _, Some(ec), _) => {
            format!("{}:{}", get_column_notation(sc), get_column_notation(ec))
        },
        // "10:18" refers to all cells in rows 10 through 18
        // "10:B18" refers to all cells in rows 10 through 18, from column B onward
        (None, Some(sr), possible_column, Some(er)) => {
            if let Some(column) = possible_column {
                // refers to all cells in given rows
                format!("{}:{}{}",  sr+1, get_column_notation(column), er+1)
            } else {
                format!("{}:{}", sr+1, er+1)
            }
        },
        _ => {
            panic!("The specified range is not valid")
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{get_a1_notation, get_column_notation};

    #[test]
    fn it_works() {
        assert_eq!(2 + 2, 4);
    }

    #[test]
    fn test_column_notation_single_letters() {
        assert_eq!(get_column_notation(0), "A");
        assert_eq!(get_column_notation(3), "D");
        assert_eq!(get_column_notation(25), "Z");
    }

    #[test]
    fn test_column_notation_double_letters() {
        assert_eq!(get_column_notation(26), "AA");
        assert_eq!(get_column_notation(27), "AB");
        assert_eq!(get_column_notation(52), "BA");
        assert_eq!(get_column_notation(701), "ZZ");
    }

    #[test]
    fn test_column_notation_high() {
        assert_eq!(get_column_notation(702), "AAA");
        assert_eq!(get_column_notation(703), "AAB");
        assert_eq!(get_column_notation(1567), "BHH");
        assert_eq!(get_column_notation(720), "AAS");
        assert_eq!(get_column_notation(14838), "UXS");
        assert_eq!(get_column_notation(17439), "YTT");
        assert_eq!(get_column_notation(18276), "ZZY");
        // highest possible column -- column ZZZ at #18277
        assert_eq!(get_column_notation(18277), "ZZZ");
    }

    #[test]
    fn test_a1_notation_001() {
        assert_eq!(
            get_a1_notation(Some(0), Some(0), Some(0), Some(0)),
            String::from("A1:A1")
        );
    }

    #[test]
    fn test_a1_notation_002() {
        assert_eq!(
            get_a1_notation(Some(0), None, Some(0), Some(0)),
            String::from("A1:A")
        );
        assert_eq!(
            get_a1_notation(Some(0), Some(0), Some(0), None),
            String::from("A1:A")
        );
    }
    #[test]
    fn test_a1_notation_003() {
        assert_eq!(
            get_a1_notation(Some(0), Some(1), Some(1), Some(4)),
            String::from("A2:B5")
        );
    }
    #[test]
    fn test_a1_notation_004() {
        assert_eq!(
            get_a1_notation(Some(0), None, Some(3), None),
            String::from("A:D")
        );
    }

    #[test]
    fn test_a1_notation_005() {
        assert_eq!(
            get_a1_notation(None, Some(9), Some(3), Some(17)),
            String::from("10:D18"),
        );

        assert_eq!(
            get_a1_notation(None, Some(9), None, Some(17)),
            String::from("10:18"),
        );
    }
}
