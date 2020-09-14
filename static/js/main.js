// Open SideNav
const toggle = document.querySelector('.toggle');
const nav = document.querySelector('.sideNav');
toggle.addEventListener('click', () => {
    toggle.classList.toggle('change');
    nav.classList.toggle('sideNav-visible');
});

// Add row to table
const addBtn = document.querySelector('.add-btn');
const tbody = document.querySelector('.tbody');
const tr = document.querySelector('tr');
addBtn.addEventListener('click', () => {
    // Get the tr tag
    const tr = document.createElement('tr');
    tr.innerHTML = `<td></td>
    <td><input type="text" id="time-hrs" placeholder="0"></td>
    <td><input type="text" id="time-min" placeholder="0"></td>
    <td><input type="text" id="height" placeholder="0"></td>
    <td><input type="text" id="gravity" placeholder="0"></td>
    <td><input type="text" id="time-hrs" placeholder="0"></td>
    <td><input type="text" id="time-min" placeholder="0"></td>
    <td><input type="text" id="height" placeholder="0"></td>
    <td><input type="text" id="gravity" placeholder="0"></td>
    <td><input type="text" id="time-hrs" placeholder="0"></td>
    <td><input type="text" id="time-min" placeholder="0"></td>
    <td><input type="text" id="height" placeholder="0"></td>
    <td><input type="text" id="gravity" placeholder="0"></td>
    <td><input type="text" id="time-hrs" placeholder="0"></td>
    <td><button type="button" class="btn btn-del">Remove</button></td>
    `;

    // Append into the tbody
    tbody.appendChild(tr);
});

// Remove ro from table
tbody.addEventListener('click', (del) => {
    if (del.target.classList.contains('btn-del')) {
        var tr = del.target.parentElement.parentElement;
        tbody.removeChild(tr);
    }
});