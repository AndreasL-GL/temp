# Website-public

## Public release
Public release of my python flask website.
Using the dockerized version to avoid displaying my normal postgres password publicly.

#### Docker-compose
Using 3 services in docker: Postgres, Nginx and Python Flask/Gunicorn

#### Website specifics:

##### Psycopg2-Functions:
Main function file, all calls to the databse using postgresql is done in this file.

### Web projects folder:
My web projects are displayed here.
##### Geocode: 
An adaption from "Python for everybody" course changed to fit my own application.
##### Getgroup: 
Annoyed with the difficulty of arranging groups I made a group compiler in python. \
Just a simple program that takes a list of names separated by commas and creates groups.

### Project features:
##### Workout-app:
A workout application i made to manage my personal workout log. \
I can use it to create workouts and add sessions for each workout. \
Because I am using high-intensity training methods I also implemented a function to give me the best calculated result for a set for each workout. This makes it very easy to know exactly what I want to do for the next session.
