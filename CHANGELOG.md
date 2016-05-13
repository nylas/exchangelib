## 1.3.6
- Don't require sequence arguments to `Folder.*_items()` methods to support `len()`
  (e.g. generators and `map` instances are now supported)
- Allow empty sequences as argument to `Folder.*_items()` methods


## 1.3.4
- Add support for `required_attendees`, `optional_attendees` and `resources`
attribute on `folders.CalendarItem`. These are implemented with a new `folders.Attendee`
class.


## 1.3.3
- Add support for `organizer` attribute on `CalendarItem`.  Implemented with a
new `folders.Mailbox` class.


## 1.2
- Initial import
