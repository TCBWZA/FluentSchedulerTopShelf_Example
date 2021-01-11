This is an example of how to use TopShelf & FluentScheduler with VB.Net.

It is coded as a VB.Net console App targeting .Net Framework 4.7.2. If you wish to target 4.5 you will need to downgrade the version of FluentScheduler to 5.3.0.

It supports:
- FluentScheduler Job events
- Multiple Jobs at the same time
- Run Once on startup with delay

The Service has startup, and stop. The pause and continue is commented out as you will need to include logic in your code to facilitate this otherwise a running job will continue to run before the pause.
