# ⚓ : Rusty 

At the time of making this repository, I found no github repositories that can provide a basic framework to test oracle forms. This project is to fill in that gap for new users.

# The Problem
I call it Rusty, since UFT in my experience is outdated and little rusty when it comes to working with new technologies  and way of working. To name a few:
- There isnt a good GIT integration. i.e. not only the integration is very basic, if I commit the code from say VS code for a function library or action, UFT wouldnt see it as committed.
- UFT still relies heavily on excels for data storage and most of its user interfacing. Excel being an application will go as a binary in GIT. You cannot collabrate with others with a binary file. 
- UFT uses object repositories for storing objects and properties for saving properties. Both these formats are non readable formats, that cannot be in any sensible way version controlled in git. If multiple people are working on it, you will only see that 'something changed' but you will have no clue 'what changed'. UFT should have come up with substitute artifacts that could be version controlled in GIT and thus make collaboration possible.
- Still no good ways to assert stuf. Yes there are checkpoints but their abilites are pretty limited. There are tons of libraries out there but its addins support seems pretty limited.
- VBScript is an old and ancient language. Although easy to learn, its very limited in todays context. 
- Still no support with parsers for parsing data types (JSON, XML, CSV, YAML). Most programming languages come with parsers to do these jobs.In UFT you have to make these yourselves.
- No standard functions for database connections. You have to make one yourself with ADODB connection and it seems many are still stuck between 32 bit and 64 bit issues. With 64 bit systems now a norm from many years, its a bit of surprise that we still have these issues.
- IT is too bulky as a tool for quick automation. It limits performance and execution speed. 
- If you rename a action, it leaves 'zomby actions' that are just 'hanging there'. 

There are probably more things that I could think of but lets leave it to that.

Nonetheless, there are some applications, which are build in ancient technologies and are still used in companies. No new tools are there around these old technologies and you may be stuck with UFT. If you are, then hopefully rusty can give you a good start with building automation frameworks.

# Rusty's Solution
Now I cannot take away some of the core limitations of UFT, but with Rusty, I have tried to take away a few of the problems to give you a better integration with GIT and thus a better chance of collaborating with other team members.
- Instead of storing objecs in object repositores (that you cant manage in GIT), we create a function library that helps identifying differnt types of objects on run time. This gives a clean and consistant way to deal with objects without adding bulky and duplicate objects. This code (or so to say virtual object repository) is git readable.
- Actions are stored in another function library. This makes the seperation of concerns (intent and implementation) possible. Giving us smaller functions that are git manageable.
- You can now combine the objects and Actions in every two lines of codes to achieve what you want to achieve. With a consistant way to build your tests, it gives a easy way to write tests (instead of writing them in excel as a keyword based approach -which git cannot work with) to a format, that is similar to the excel one but also git maintainable).See example files to understand what I mean here.
- Excels, although not used at the time of writing, will be replaced with either XMLs, or CSVs, both these artifacts being git friendly.
- In general, all artifacts are created keeping a clean design, git maintanability, collaboration and long term maintanece in mind. 

# Scope
    - Oracle Forms
    - Web browsers
    - Files

# [Naming conventions](https://medium.com/better-programming/string-case-styles-camel-pascal-snake-and-kebab-case-981407998841)
* Naming directories and files
    * Directory names: PascalCase
    * File names: kebab-case
    * No spaces allowed between Directory and File names (So "Unified Functional Tester" and "Assert function library.txt" are bad directory and file names)
* [Naming functions, parameters and variables](https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/program-structure/naming-conventions)
    * Function Names: PascalCase 
    * Parameters: camelCase
    * Variables: camelCase
    * Constants: SNAKE_CASE
    * Database fields: snake_case
    * urls: kebab-case