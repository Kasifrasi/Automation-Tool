//! Cell address type for the report grid.

use std::fmt;

/// Excel cell address (0-based internally).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct CellAddr {
    pub row: u32,
    pub col: u16,
}

impl CellAddr {
    pub const fn new(row: u32, col: u16) -> Self {
        Self { row, col }
    }

    /// From Excel notation: A=0, B=1, ..., row is 1-based in Excel.
    /// E.g. E2 = (row=1, col=4)
    pub const fn excel(col: u16, row: u32) -> Self {
        Self { row: row - 1, col }
    }

    /// Converts to Excel notation (e.g. "E2").
    pub fn to_excel(self) -> String {
        let col_letter = Self::col_to_letter(self.col);
        format!("{}{}", col_letter, self.row + 1)
    }

    pub(crate) fn col_to_letter(col: u16) -> String {
        let mut result = String::new();
        let mut c = col as u32;
        loop {
            result.insert(0, (b'A' + (c % 26) as u8) as char);
            if c < 26 {
                break;
            }
            c = c / 26 - 1;
        }
        result
    }
}

impl fmt::Display for CellAddr {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{}", self.to_excel())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_addr_to_excel() {
        assert_eq!(CellAddr::new(0, 0).to_excel(), "A1");
        assert_eq!(CellAddr::new(1, 4).to_excel(), "E2");
        assert_eq!(CellAddr::new(30, 21).to_excel(), "V31");
    }
}
