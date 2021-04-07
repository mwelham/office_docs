# Brief

Provide access to internal structure of docx and xlsx files.

Provide placeholder templating for content.

## Tests

Most of the `docx` tests are in minispec. All of the `xlsx` tests and some
`docx` tests are in rspec.

To run all tests say
```
rake default
```

### minispec

Normal minispec (ie not the slower tests..?):
```
rake test
```

All minispec tests (the arabic date/time one is failing as of 07-Apr-2021):
```
rake test:all
```
### rspec

Run normal rspec tests:
```
rspec
```

There are some rspec tags:

```
rspec -t performance # shows some bmbm comparisons
rspec -t extracted   # Placeholder parsing of all placeholder strings used by the rake tests.
rspec -t display_ui  # shows some rendered templates
rspec -t all         # everything except display_ui
```

Note that `spec_helper.rb` will automatically run a check that the parser has been built from the grammar file. To disable this check (which takes little time anyway) set env var `NO_RAKE_GRAMMAR=true`,

for example:
```
NO_RAKE_GRAMMAR=true rspec spec/sheet_spec.rb
```
