const tenant = "tenant";
const client_id = 'clientid';
const client_secret = 'clientsecret';

const graph = require("@microsoft/microsoft-graph-client");
const express = require('express');
const session = require('express-session')
const request = require('request');
const https = require('https');
const http = require('http');
const { query } = require("express");

var app = express();
app.use('/static', express.static('public'));
app.use(session({
    secret: '',
    cookie: {}
}))

var authorize_uri = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/authorize`;
var token_uri = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
var redirect_uri = 'http://localhost:8080/authorize';

var client_scopes = 'User.Read User.ReadBasic.All offline_access';

// We have 2 types of requests to make to the v2 endpoint, first to get the access token and second to get the refresh token
var token_request = {
    form: {
        grant_type: 'authorization_code',
        code: '',
        client_id: client_id,
        client_secret: client_secret,
        scope: client_scopes,
        redirect_uri: redirect_uri
    }
}

var refresh_token_request = {
    form: {
        grant_type: 'refresh_token',
        refresh_token: '',
        client_id: client_id,
        client_secret: client_secret,
        scope: client_scopes,
        redirect_uri: redirect_uri
    }
}

// This is the web root and provides a link that kicks off the OAUTH process 
app.get('/', function(req, res) {
    var codegrant_endpoint = authorize_uri + '?client_id=' + client_id + '&response_type=code&redirect_uri=' + redirect_uri + '&scope=' + client_scopes + '&response_mode=query&state=12345&nonce=678910';
    res.send('<div><a href="' + codegrant_endpoint + '" target="_blank">GET ACCESS TOKEN</a></div>');
});

/*
    Here we convert the auth code returned by the endpoint into a bearer token we can use to call the API.
*/
app.get('/authorize', function(req, res) {
    if(req.query.error != null)
    {        
        var content = "<p>";
        content += req.query.error;
        content += "</p>";
        content += "<p>";
        content += req.query.error_description;
        content += "</p>";
        res.end(content)
    }
    else if (req.query.code != null) {
        // get the auth code from the query params
        var auth_code = req.query.code;

        // Add the code in our token_request
        token_request.form.code = auth_code;
  
        request.post(token_uri, token_request, function(err, httpResponse, body) {
            var result = JSON.parse(body);

            var content = "<pre>" + JSON.stringify(result, null, 2) + "</pre>"
            content += "</p>";

            getProfile(result.access_token, function(profile) {
                content += "<pre>" + JSON.stringify(profile, null, 2) + "</pre>";
                content += "</p>";

                content += '<a href="/refresh?code=' + result.refresh_token + '" target="">GET REFRESH TOKEN</a>';
                res.end(content)
            });
        })
    } else {
        // OAUTH Implicit Grant workflow

        var token = {
            access_token: req.query.access_token,
            token_type: req.query.token_type,
            expires_in: req.query.expires_in,
            scope: req.query.scope
        }

        var content = "<pre>" + JSON.stringify(token, null, 2) + "</pre>"
        content += "<pre>" + token.access_token + "</pre>";
        content += '<a href="/refresh?code=' + result.refresh_token + '" target="">GET REFRESH TOKEN</a>'
        res.end(content)

    }
});

// To make the refresh token into a usable bearer token
app.get('/refresh', function(req, res) {

    // get the refresh_token from the query params
    var refresh_token = req.query.code;

    // Add the code in our refresh_token_request
    refresh_token_request.form.refresh_token = refresh_token;

    request.post(token_uri, refresh_token_request, function(err, httpResponse, body) {
        var result = JSON.parse(body);

        var content = "<pre>" + JSON.stringify(result, null, 2) + "</pre>"
        content += "<pre>" + refresh_token + "</pre>";
        content += '<a href="/refresh?code=' + result.refresh_token + '" target="">GET REFRESH TOKEN</a>'
        res.end(content)
    })
});

function getProfile(access_token, callback) {
    var client = graph.Client.init({
        authProvider: (done) => {
            done(null, access_token);
        }
    });
    client
        .api('/me')
        .get((err, res) => {
            callback(res);
        });
}

http.createServer(app).listen(8080);

console.log('App listening on port 8080');