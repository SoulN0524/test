from flask import Flask, request, jsonify
app=Flask(__name__)

students=[{
    "name":"Joke",
    "age":48,
    "courses":["MATH","STATISTICS"]
},
{
    "name":"Sam",
    "age":59,
    "courses":["HISTORY","PROGRAM"]
}]
@app.route("/student",methods=["GET"])
def get_students():
    return jsonify({"students":students})

@app.route("/student", methods=["POST"])
def create_students():
    request_data=request.get_json()
    new_student={"name":request_data["name"],"age":request_data["age"],"courses":[]}
    students.append(new_student)
    return jsonify(new_student),201
@app.route("/student/<string:name>/course",methods=["POST"])
def add_course_to_student(name):
    request_data=request.get_json()
    for student in students:
        if student["name"] ==name:
            new_course=request_data["course"]
            student["courses"].append(new_course)
            return jsonify({"message": f"Course {new_course} added to student {name}"}), 201
    return jsonify({"message": "Student not found"}), 404

if __name__=="__main__":
    app.run(debug=True)
