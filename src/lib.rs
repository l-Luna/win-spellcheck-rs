use std::sync::Mutex;
use windows::{Win32::Foundation::*, Win32::Globalization::*, Win32::System::Com::*, core::*};

/// A simple wrapper around the Windows Spellchecking API.
///
/// Create a [Spellchecker] of the desired locale, then call [Spellchecker::check] to check a string.
/// The returned [SpellingError]s contain replacement suggestions.

// try to only initialize COM once
// (though it's OK if this happens multiple times or other libraries do so too)
static COM_INIT: Mutex<bool> = Mutex::new(false);

fn try_init_com() {
    let mut com_init = COM_INIT.lock().unwrap();
    if !*com_init {
        *com_init = true;
        drop(com_init);
        unsafe {
            CoInitializeEx(None, COINIT_MULTITHREADED)
                .ok()
                .expect("Failed to initialize COM!");
        }
    }
}

pub struct Spellchecker(ISpellChecker);

#[derive(Clone, Debug, Eq, PartialEq, Ord, PartialOrd, Hash, Default)]
pub enum Correction {
    #[default]
    None,
    Delete,
    Suggestions(Vec<String>),
    Replacement(String),
}

#[derive(Clone, Debug, Eq, PartialEq, Ord, PartialOrd, Hash, Default)]
pub struct SpellingError {
    pub start: usize,
    pub length: usize,
    pub correction: Correction,
}

impl Spellchecker {
    pub fn new(locale: &str) -> Option<Self> {
        try_init_com();
        let factory: ISpellCheckerFactory =
            unsafe { CoCreateInstance(&SpellCheckerFactory, None, CLSCTX_ALL) }.ok()?;
        let locale = HSTRING::from(locale);
        let local_supported = unsafe { factory.IsSupported(&locale) }.ok()?;
        if !local_supported.as_bool() {
            return None;
        }
        let checker = unsafe { factory.CreateSpellChecker(&locale) }.ok()?;
        Some(Self(checker))
    }

    pub fn new_en() -> Option<Self> {
        Self::new("en-US")
    }

    pub fn check(&self, text: &str) -> Option<Vec<SpellingError>> {
        let errors = unsafe { self.0.ComprehensiveCheck(&HSTRING::from(text)) }.ok()?;
        let mut err = None;
        let mut results = Vec::new();
        while unsafe { errors.Next(&mut err) } == S_OK {
            let err = err.take().unwrap();
            let start = unsafe { err.StartIndex() }.ok()? as usize;
            let length = unsafe { err.Length() }.ok()? as usize;
            let correction = unsafe { err.CorrectiveAction() }.ok()?;
            let correction: Correction = match correction {
                CORRECTIVE_ACTION_DELETE => Correction::Delete,
                CORRECTIVE_ACTION_GET_SUGGESTIONS => {
                    let mut results = Vec::new();
                    let substring = &text[start..(start + length)];
                    let suggestions = unsafe { self.0.Suggest(&HSTRING::from(substring)) }.ok()?;
                    let mut suggestion = [PWSTR::null()];
                    while unsafe { suggestions.Next(&mut suggestion, None) } == S_OK
                        && !suggestion[0].is_null()
                    {
                        results.push(unsafe { suggestion[0].to_string() }.ok()?);
                        unsafe { CoTaskMemFree(Some(suggestion[0].as_ptr() as *mut _)) };
                    }

                    Correction::Suggestions(results)
                }
                CORRECTIVE_ACTION_REPLACE => {
                    let replacement = unsafe { err.Replacement() }.ok()?;
                    let replacement_s = unsafe { replacement.to_string() }.ok()?;
                    unsafe { CoTaskMemFree(Some(replacement.as_ptr() as *mut _)) };
                    Correction::Replacement(replacement_s)
                }
                _ => Correction::None,
            };
            results.push(SpellingError {
                start,
                length,
                correction,
            });
        }
        Some(results)
    }
}

#[cfg(test)]
mod tests {
    use crate::Spellchecker;

    #[test]
    fn test() {
        let spellchecker = Spellchecker::new_en().expect("Failed to create english spellchecker!");
        let text = "another one bitess the dust; another whitness blinded";
        let errors = spellchecker.check(text).unwrap();
        println!("{errors:?}");

        let spellchecker2 = Spellchecker::new_en().unwrap();
        let errors2 = spellchecker2.check(text).unwrap();
        assert_eq!(errors, errors2);
    }
}
