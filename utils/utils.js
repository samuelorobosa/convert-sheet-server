function stringToUTC(year, month, day, hour = 6, minute = 30, second = 0){
    const date = new Date(Date.UTC(year, month - 1, day, hour, minute, second));
    return  date.toISOString().replace(/-|:|\.\d+/g, "");
}



module.exports = {stringToUTC};