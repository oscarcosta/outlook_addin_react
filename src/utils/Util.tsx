
export default class Util {
    static log = (text) => {
        console.log(text)
        if (document && document.getElementById("log")) {
            document.getElementById("log").innerText += " \n " + JSON.stringify(text)
        }
    }
}
