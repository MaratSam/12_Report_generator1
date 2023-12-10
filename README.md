# Locating-Report-Generator
Copyright Marat Samigullin 2023

**This Python Windows TKinter desktop app creates a MS Word template of a service locating report according to parameters which user inputs whrough a dialog window,
such as address, docket number, parameters of service locating job, date etc.**
The name of the created file consists of the address and suburb name.
It also populates the report with photographs.
The user should prepare the photorgaphs by themseves separately in a image processing app like paint.net or similar, as well as a map of the service locating area.
All pre-setup phrases are stored in templates.json file.

The app also creates a template for the email to be sent to the client (file MS Word email template for the client) to streamline the process of emailng the report.

To adopt the project to your own needs, it is recommended to edit Title_image_1.jpg and Title_image_2.jpg as well as templates.json
