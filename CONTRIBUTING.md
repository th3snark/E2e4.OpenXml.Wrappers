# Contributing

Thanks for helping. Before opening a pull request:

```bash
dotnet build E2e4.OpenXml.Wrappers.slnx
dotnet test E2e4.OpenXml.Wrappers.slnx
dotnet format E2e4.OpenXml.Wrappers.slnx --verify-no-changes
```

- The build must have zero warnings (`CI=true` makes warnings errors, as in CI).
- Follow `.editorconfig` and the conventions in [CLAUDE.md](CLAUDE.md): explicit types, 0-based indexes, a test class named `<Class>Test` for each wrapper.
- New behaviour needs tests. If it writes new XML, keep the schema element order and make sure the result passes `OpenXmlValidator` and opens in Word.
- New examples go in `samples/OpenXml.Wrappers.Samples` and in `Examples.All`; the samples test validates them.
- Update the README and `CHANGELOG.md` for any change to the public API.
