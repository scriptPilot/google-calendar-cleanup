// This function resets the script properties (e.g., if a script gets stuck in a waiting state)
function resetScript() {
  PropertiesService.getUserProperties().deleteAllProperties()  
  console.log('Script reset done.')
}
