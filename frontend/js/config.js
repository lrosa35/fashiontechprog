// Auto-pick API endpoint based on host
(function(){
  var isLocal = /^(localhost|127\.0\.0\.1)$/i.test(window.location.hostname);
  window.API_BASE = isLocal ? 'http://localhost:8090' : 'https://fashiontechprog-255445129a01\.herokuapp\.com';
})();
