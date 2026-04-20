# Changelog

All notable changes to Master Register are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-04-19

### Added
- Initial release of Master Register script
- Automatic monthly sheet creation from existing template
- Configurable column and row mappings for any spreadsheet layout
- Automatic day column adjustment for different month lengths
- Format preservation when duplicating sheets
- Duplicate sheet prevention (won't create if sheet already exists)
- Interactive menu system ("Register" menu in Google Sheets)
- Comprehensive configuration system (CFG object)
- Extensive error handling with user-friendly messages
- Email validation and helpful error messages for common issues

### Features
- One-click sheet creation from spreadsheet menu
- Automatic month/year calculation
- Smart day column management (adds/removes columns as needed)
- Format and formula preservation
- Automatic clearing of attendance data
- Prevention of duplicate sheet creation
- Validation of configuration settings

### Documentation
- README.md with complete feature overview
- SETUP_GUIDE.md with step-by-step installation
- EXAMPLES.md with configuration examples for different layouts
- TROUBLESHOOTING.md with common issues and solutions
- CHANGELOG.md (this file) with version history

### Configuration
- Six configurable settings to match any spreadsheet layout:
  - monthCell: Location of month name
  - yearCell: Location of year value
  - dayNumbersRow: Row containing day numbers
  - firstDayCol: First lesson date column
  - firstPupilRow: First student row
  - lastPupilRow: Last student row

### Testing
- Tested with multiple spreadsheet layouts
- Verified compatibility with different month lengths (28-31 days)
- Confirmed proper handling of edge cases (year transitions, etc.)

---

## Future Roadmap

### Potential Features for Future Versions
- Automatic schedule push to calendar services
- Email notifications for sheet creation
- Batch creation (create multiple months at once)
- Sheet archival automation
- Customizable template per month
- Integration with Google Calendar
- Pupil progress tracking enhancements

### Potential Improvements
- Enhanced error messages with suggested fixes
- Configuration wizard to auto-detect settings
- Sheet rename/reorganization features
- Backup functionality for archives
- Analytics and reporting features
- Mobile app integration

---

## Support

For issues, feature requests, or suggestions:
1. Review the troubleshooting guide
2. Check example configurations
3. Verify your setup against the setup guide
4. Check common issues in the troubleshooting documentation

---

## Credits

**Master Register** was created to automate the monthly administration tasks for Piano Lessons with Greg Kaighin.

### Technologies Used
- Google Apps Script
- Google Sheets API
- JavaScript

---

## License

This project is licensed under the MIT License - see [LICENSE](LICENSE) file for details.

---

## Versioning

This project uses Semantic Versioning:
- MAJOR version when incompatible API changes are made
- MINOR version when functionality is added in a backwards-compatible manner
- PATCH version when backwards-compatible bug fixes are released

---

### Version Format
- `[1.0.0]` - Release version
- `[Unreleased]` - Development version with changes not yet released

### Date Format
- ISO 8601 format: `YYYY-MM-DD`

---

Last updated: 2026-04-19
