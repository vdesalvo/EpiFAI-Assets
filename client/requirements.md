## Packages
clsx | Utility for constructing className strings conditionally
tailwind-merge | Utility for merging Tailwind CSS classes
class-variance-authority | For creating reusable component variants
lucide-react | Icon library
@radix-ui/react-slot | For polymorphic components
@radix-ui/react-tabs | For tabbed interface
@radix-ui/react-dialog | For modals/dialogs
@radix-ui/react-toast | For notifications
@radix-ui/react-label | For form labels
@radix-ui/react-switch | For toggle switches
@radix-ui/react-scroll-area | For custom scrollbars
framer-motion | For smooth animations

## Notes
Must add <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script> to index.html manually if not present.
The app relies on the global `Excel` namespace provided by office.js.
Development requires running in an environment where Office.js can initialize (or mocking it).
