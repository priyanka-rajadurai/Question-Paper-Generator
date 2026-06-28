let count=0;

function addPart(){

count++;

let div=document.createElement("div");

div.className="part";

div.innerHTML=`

<h3>Part ${count}</h3>

<input name="part_name[]" placeholder="Part Name" required>

<input name="part_marks[]" placeholder="Marks" required>

<input name="part_co[]" placeholder="CO Type" required>

<input name="part_questions[]" placeholder="Number of Questions" required>

<input name="part_description[]" placeholder="Add Description">

<select name="optional[]">

<option value="NO">

NO

</option>

<option value="YES">

YES

</option>

</select>

`;

document.getElementById("parts").appendChild(div);

}