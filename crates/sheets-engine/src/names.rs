//! Defined names (Formulas → Name Manager): `Sales` for `Sheet1!$B$2:$B$100`,
//! `GST` for `0.18`. Formulas use them like references or values
//! (`=SUM(Sales)*(1+GST)`). A name belongs to the workbook, or to one sheet
//! (`scope`), which then wins over a workbook name of the same spelling there.
//!
//! The workbook keeps the list; `SetNames` replaces it as one undoable step.
//! Formulas that use names are recalculated when what a name refers to
//! changes (see `Workbook::link_index`).

use office_core::Op;

use crate::cell::CellRef;
use crate::formula::{self, Expr};
use crate::workbook::Workbook;

#[derive(Debug, Clone, PartialEq)]
pub struct DefinedName {
    pub name: String,
    /// What it refers to, as formula text without `=`: `Sheet1!$B$2:$B$100`, `0.18`.
    pub refers_to: String,
    /// `Some(i)`: only on sheet `i`; `None`: the whole workbook.
    pub scope: Option<usize>,
    pub comment: String,
    expr: Expr,
}

impl DefinedName {
    pub fn new(name: &str, refers_to: &str, scope: Option<usize>) -> DefinedName {
        let refers_to = refers_to.trim().trim_start_matches('=').to_string();
        let expr = formula::parse(&refers_to).unwrap_or_else(Expr::Error);
        DefinedName { name: name.trim().to_string(), refers_to, scope, comment: String::new(), expr }
    }

    pub fn expr(&self) -> &Expr {
        &self.expr
    }

    /// The same name referring to something else.
    pub fn with_refers_to(&self, refers_to: &str) -> DefinedName {
        DefinedName { comment: self.comment.clone(), ..DefinedName::new(&self.name, refers_to, self.scope) }
    }
}

/// Why a name can't be used, in the words Excel's dialog uses.
pub fn check_name(name: &str) -> Result<(), &'static str> {
    let name = name.trim();
    let mut chars = name.chars();
    let Some(first) = chars.next() else { return Err("Enter a name.") };
    if name.chars().count() > 255 {
        return Err("A name can be at most 255 characters long.");
    }
    if !(first.is_alphabetic() || first == '_' || first == '\\') {
        return Err("A name must start with a letter, an underscore (_) or a backslash (\\).");
    }
    if !chars.all(|c| c.is_alphanumeric() || matches!(c, '_' | '.' | '\\' | '?')) {
        return Err("A name can't contain spaces or symbols other than _ . \\ ?");
    }
    let upper = name.to_ascii_uppercase();
    if upper == "TRUE" || upper == "FALSE" || upper == "R" || upper == "C" {
        return Err("That word is reserved.");
    }
    // Anything that reads as a cell reference (A1, XFD1048576, R1C1) is ambiguous.
    if CellRef::parse(name).is_some() || looks_like_r1c1(&upper) {
        return Err("A name can't look like a cell reference, such as A1 or R1C1.");
    }
    Ok(())
}

fn looks_like_r1c1(upper: &str) -> bool {
    let Some(rest) = upper.strip_prefix('R') else { return false };
    let digits_r: String = rest.chars().take_while(|c| c.is_ascii_digit()).collect();
    let rest = &rest[digits_r.len()..];
    let Some(rest) = rest.strip_prefix('C') else { return rest.is_empty() && !digits_r.is_empty() };
    rest.chars().all(|c| c.is_ascii_digit())
}

impl Workbook {
    /// The name `name` as seen from sheet `sheet`: a sheet-scoped name first, then a workbook one.
    pub fn defined_name(&self, name: &str, sheet: usize) -> Option<&DefinedName> {
        let same = |d: &&DefinedName| d.name.eq_ignore_ascii_case(name);
        self.names
            .iter()
            .filter(same)
            .find(|d| d.scope == Some(sheet))
            .or_else(|| self.names.iter().filter(same).find(|d| d.scope.is_none()))
    }

    /// The list with `name` added or replaced (same name and scope). Errors use Excel's wording.
    pub fn with_name(
        &self,
        name: DefinedName,
        replacing: Option<(&str, Option<usize>)>,
    ) -> Result<Vec<DefinedName>, String> {
        check_name(&name.name).map_err(String::from)?;
        if matches!(name.expr, Expr::Error(_)) {
            return Err("The reference isn't valid. Use a range such as Sheet1!$A$1:$A$10, or a value.".into());
        }
        let mut list: Vec<DefinedName> = self
            .names
            .iter()
            .filter(|d| match replacing {
                Some((n, scope)) => !(d.name.eq_ignore_ascii_case(n) && d.scope == scope),
                None => true,
            })
            .cloned()
            .collect();
        if list.iter().any(|d| d.name.eq_ignore_ascii_case(&name.name) && d.scope == name.scope) {
            return Err("That name already exists. Names must be unique.".into());
        }
        list.push(name);
        list.sort_by(|a, b| a.name.to_lowercase().cmp(&b.name.to_lowercase()).then(a.scope.cmp(&b.scope)));
        Ok(list)
    }

    /// The names with every `refers_to` passed through `rewrite` (which takes and
    /// returns formula text with its `=`), e.g. after a sheet was renamed.
    pub fn names_rewritten(&self, rewrite: impl Fn(&str) -> String) -> Vec<DefinedName> {
        self.names
            .iter()
            .map(|d| {
                let text = rewrite(&format!("={}", d.refers_to));
                d.with_refers_to(text.trim_start_matches('='))
            })
            .collect()
    }

    /// Recalculate every formula that uses a name (after the names changed).
    pub(crate) fn recalc_name_users(&mut self) {
        let changed: Vec<(usize, CellRef)> =
            self.sheets.iter().enumerate().flat_map(|(s, ws)| ws.sheet.name_users().map(move |c| (s, c))).collect();
        if !changed.is_empty() {
            self.recalc_cells(changed);
        }
    }
}

/// Replace the workbook's defined names.
#[derive(Debug, Clone)]
pub struct SetNames {
    pub names: Vec<DefinedName>,
}

impl Op<Workbook> for SetNames {
    fn apply(&self, doc: &mut Workbook) {
        doc.names = self.names.clone();
        doc.recalc_name_users();
    }

    fn invert(&self, doc: &Workbook) -> Box<dyn Op<Workbook>> {
        Box::new(SetNames { names: doc.names.clone() })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Value;

    fn at(s: &str) -> CellRef {
        CellRef::parse(s).unwrap()
    }

    #[test]
    fn names_follow_excels_rules() {
        assert!(check_name("Sales_2025").is_ok());
        assert!(check_name("_tax").is_ok());
        assert!(check_name("A1").is_err());
        assert!(check_name("R1C1").is_err());
        assert!(check_name("my name").is_err());
        assert!(check_name("1st").is_err());
        assert!(check_name("TRUE").is_err());
    }

    #[test]
    fn formulas_use_names_and_follow_their_changes() {
        let mut wb = Workbook::new();
        wb.set_inputs(0, [(at("B2"), "100".into()), (at("B3"), "50".into()), (at("D1"), "=SUM(Sales)*(1+GST)".into())]);
        assert_eq!(wb.sheets[0].sheet.value(at("D1")), &Value::Error(crate::ErrorKind::Name));
        let names = wb.with_name(DefinedName::new("Sales", "Sheet1!$B$2:$B$3", None), None).unwrap();
        let names = Workbook { names, ..wb.clone() }.with_name(DefinedName::new("GST", "0.18", None), None).unwrap();
        SetNames { names }.apply(&mut wb);
        assert_eq!(wb.sheets[0].sheet.value(at("D1")), &Value::Number(177.0));
        // Changing a cell the name covers recalculates the formula.
        wb.set_inputs(0, [(at("B3"), "0".into())]);
        assert_eq!(wb.sheets[0].sheet.value(at("D1")), &Value::Number(118.0));
        // Duplicates and bad references are refused.
        assert!(wb.with_name(DefinedName::new("sales", "Sheet1!A1", None), None).is_err());
        assert!(wb.with_name(DefinedName::new("X", "Sheet1!A1:", None), None).is_err());
    }

    #[test]
    fn names_follow_sheet_renames_and_deletes() {
        use crate::ops::{RemoveSheet, RenameSheet};
        let mut wb = Workbook::new();
        wb.add_sheet();
        SetNames {
            names: vec![DefinedName::new("Data", "Sheet2!$A$1:$A$3", None), DefinedName::new("Local", "1", Some(1))],
        }
        .apply(&mut wb);
        RenameSheet { index: 1, name: "Sales Data".into() }.apply(&mut wb);
        assert_eq!(wb.names[0].refers_to, "'Sales Data'!$A$1:$A$3");
        let undo = RemoveSheet { index: 1 }.invert(&wb);
        RemoveSheet { index: 1 }.apply(&mut wb);
        assert_eq!(wb.names.len(), 1);
        assert_eq!(wb.names[0].refers_to, "#REF!");
        undo.apply(&mut wb);
        assert_eq!(wb.names.len(), 2);
        assert_eq!(wb.names[0].refers_to, "'Sales Data'!$A$1:$A$3");
        assert_eq!(wb.names[1].scope, Some(1));
    }

    #[test]
    fn sheet_names_win_over_workbook_names() {
        let mut wb = Workbook::new();
        wb.add_sheet();
        let names = vec![DefinedName::new("Rate", "5", None), DefinedName::new("Rate", "7", Some(1))];
        SetNames { names }.apply(&mut wb);
        wb.set_inputs(0, [(at("A1"), "=Rate".into())]);
        wb.set_inputs(1, [(at("A1"), "=Rate".into())]);
        assert_eq!(wb.sheets[0].sheet.value(at("A1")), &Value::Number(5.0));
        assert_eq!(wb.sheets[1].sheet.value(at("A1")), &Value::Number(7.0));
    }
}
