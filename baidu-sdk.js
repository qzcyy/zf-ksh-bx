var axios = require('axios')
var qs = require('qs')

module.exports = {
    getAuth: function (api_key, secret_key) {
        return axios.post("https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=" + api_key + "&client_secret=" + secret_key)
    },
    getImageDetail: function (access_token, body) {
        return axios.post("https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic?access_token=" + access_token,qs.stringify(body))
    }
}