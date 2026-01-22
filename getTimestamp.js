// Get timestamp
function getTimestamp() {
  var today = new Date();
  var yyyy = today.getFullYear();
  var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
  var dd = String(today.getDate()).padStart(2, '0');
  var hh = String(today.getHours()).padStart(2, '0');
  var mn = String(today.getMinutes()).padStart(2, '0');
  var sc = String(today.getSeconds()).padStart(2, '0');

  return yyyy + mm + dd + '-' + hh + mn + sc;
}
