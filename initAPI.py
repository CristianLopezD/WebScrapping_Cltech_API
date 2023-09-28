from datetime import datetime, date
from flask import Flask, request, jsonify, send_file
import os
import Logic 

app = Flask(__name__)

# Define a route for your endpoint
@app.route('/getCSVFile', methods=['GET'])
def getCSVFile():
    try:
        # Get the parameters from the request JSON
        data = request.json
        TicketsToSearch = data.get('Tickets')
        Key = data.get('Key')

        # Validate the Key format
        try:
            key_date = datetime.strptime(Key, 'Cltech%Y%m%d')
        except ValueError:
            response = {'message': 'Error', 'error': 'Invalid Key format'}
            return jsonify(response), 400

        # Compare the key_date to the current date
        if key_date.date() != date.today():
            response = {'message': 'Error', 'error': 'Invalid Key date'}
            return jsonify(response), 400
        
        if os.path.exists("ticket_data.csv"):
            os.remove("ticket_data.csv")

        Logic.init_APP(TicketsToSearch)
        if os.path.exists("ticket_data.csv"):
            response = send_file("ticket_data.csv", as_attachment=True)
            return response
        else:
            # If the file doesn't exist, return an error response
            response = {'message': 'Error', 'error': 'ticket_data.csv not found'}
            return jsonify(response), 404

    except Exception as e:
        # Handle any errors and return an error response
        error_message = str(e)
        response = {'message': 'Error', 'error': error_message}
        return jsonify(response), 500

@app.route('/helloWorld', methods=['GET'])
def helloWorld():
    try:
        response = {'message': 'Good', 'Message': 'Hello '}
        return jsonify(response), 404

    except Exception as e:
        # Handle any errors and return an error response
        error_message = str(e)
        response = {'message': 'Error', 'error': error_message}
        return jsonify(response), 500

if __name__ == '__main__':
    # Run the Flask app on a public IP address and port (you can change these values)
    app.run(host='0.0.0.0', port=5000)
