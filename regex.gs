function mySearchFunction() {
  var searchVerb = "log {\"";
  var searchTerm = "log *{\"";
  //Logger.log(searchTerm.test("log *{\""));
  Logger.log(searchVerb.search(searchTerm));
  searchTerm = "shield_[ssl]*cache";
  searchVerb = "shield_ssl_cache";
  Logger.log(searchVerb.search(searchTerm));
}
