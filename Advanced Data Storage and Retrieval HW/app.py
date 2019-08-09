# Import Dependencies 

import numpy as np

import sqlalchemy
from sqlalchemy.ext.automap import automap_base
from sqlalchemy.orm import Session
from sqlalchemy import create_engine, func
import datetime as dt
import pandas as pd

from flask import Flask, jsonify


#################################################
# Database Setup
#################################################
engine = create_engine("sqlite:///Resources/hawaii.sqlite")

# reflect the database into a new model
Base = automap_base()
# reflect the tables
Base.prepare(engine, reflect=True)
print(Base.classes.keys())
# Save reference to the table
Measurement = Base.classes.measurement
Station = Base.classes.station


# Create our session (link) from Python to the DB
session = Session(engine)

#################################################
# Flask Setup
#################################################
app = Flask(__name__)

#################################################
# Flask Routes
#################################################

@app.route("/")
def welcome():
    """List all available api routes."""
    return (
        f"Available Routes:<br/>"
        f"/api/v1.0/precipitation<br/>"
        f"/api/v1.0/stations<br/>"
        f"/api/v1.0/tobs<br/>"
        f"/api/v1.0/start<br/>"
        f"/api/v1.0/start/end"
    )

@app.route("/api/v1.0/precipitation")
def precipitation():
    """Return a list of rain fall for prior year"""
    # Query for the dates and precipitation observations from the last year.
    # Convert the query results to a Dictionary using `date` as the key and `prcp` as the value.
    # Return the JSON representation of your dictionary.
    
    recent_date = session.query(Measurement.date).order_by(Measurement.date.desc()).first()
    one_year_ago = dt.date(2017, 8, 23) - dt.timedelta(days=365)
    rain = session.query(Measurement.date, Measurement.prcp).\
        filter(Measurement.date > one_year_ago).\
        order_by(Measurement.date).all()

    # Create a list of dicts with `date` and `prcp` as the keys and values
    
    rain_totals = []
    for result in rain:
        row = {}
        row["date"] = rain[0]
        row["prcp"] = rain[1]
        rain_totals.append(row)

    return jsonify(rain_totals)

@app.route("/api/v1.0/stations")
def stations():
    # Return a JSON list of stations from the dataset.
    stations_query = session.query(Station.name, Station.station)
    stations = pd.read_sql(stations_query.statement, stations_query.session.bind)
    
    return jsonify(stations.to_dict())


@app.route("/api/v1.0/tobs")
def tobs():
    """Return a list of temperatures for prior year"""
    
    # Query for the dates and temperature observations from a year from the last data point.
    # Return a JSON list of Temperature Observations (tobs) for the previous year.
    
    recent_date = session.query(Measurement.date).order_by(Measurement.date.desc()).first()
    one_year_ago = dt.date(2017, 8, 23) - dt.timedelta(days=365)
    temperature = session.query(Measurement.date, Measurement.tobs).\
        filter(Measurement.date > one_year_ago).\
        order_by(Measurement.date).all()

    # Create a list of dicts with `date` and `tobs` as the keys and values
    
    temperature_totals = []
    for result in temperature:
        row = {}
        row["date"] = temperature[0]
        row["tobs"] = temperature[1]
        temperature_totals.append(row)

    return jsonify(temperature_totals)

# Return a JSON list of the minimum temperature, the average temperature, and the max temperature for a given start or start-end range.

# When given the start only, calculate `TMIN`, `TAVG`, and `TMAX` for all dates greater than and equal to the start date.

# When given the start and the end date, calculate the `TMIN`, `TAVG`, and `TMAX` for dates between the start and end date inclusive.

@app.route("/api/v1.0/<start>")
def trip1(start):
 
    start_date= dt.datetime.strptime(start, '%Y-%m-%d')
    one_year_ago = dt.timedelta(days=365)
    start = start_date-one_year_ago
    end =  dt.date(2017, 8, 23)
    trip_data = session.query(func.min(Measurement.tobs), func.avg(Measurement.tobs), func.max(Measurement.tobs)).\
        filter(Measurement.date >= start).filter(Measurement.date <= end).all()
    trip = list(np.ravel(trip_data))
    
    return jsonify(trip)


@app.route("/api/v1.0/<start>/<end>")
def trip2(start,end):
    
    start_date= dt.datetime.strptime(start, '%Y-%m-%d')
    end_date= dt.datetime.strptime(end,'%Y-%m-%d')
    one_year_ago = dt.timedelta(days=365)
    start = start_date-one_year_ago
    end = end_date-one_year_ago
    trip_data = session.query(func.min(Measurement.tobs), func.avg(Measurement.tobs), func.max(Measurement.tobs)).\
        filter(Measurement.date >= start).filter(Measurement.date <= end).all()
        
    trip = list(np.ravel(trip_data))
    
    return jsonify(trip)



if __name__ == "__main__":
    app.run(debug=True)
    