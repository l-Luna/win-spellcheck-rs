a simple wrapper around the Windows Spellchecking API

usage:
```rs
let spellchecker = Spellchecker::new_en().expect("Failed to create english spellchecker!");
let text = "another one bitess the dust; another whitness blinded";
let errors = spellchecker.check(text).unwrap();
println!("{errors:?}");

// prints: [SpellingError { start: 12, length: 6, correction: Replacement("bites") }, SpellingError { start: 37, length: 8, correction: Suggestions(["whatness", "whiteness", "witness", "whitens", "whites", "whines"]) }]
```