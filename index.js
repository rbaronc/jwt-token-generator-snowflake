const crypto = require('crypto');
const fs = require('fs');
const jwt = require('jsonwebtoken');
const express = require('express');

const app = express();
const port = 3000;
const qualifiedUsername = "OXVLXZE-QIB33438.RBARONC"; //Modify this if you try to use another account. You need to configure the public key in snowflake if you do that .
const privateKeyFilePath = './rsa_key.p8';


app.use(function(_, res, next) {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Credentials', true);
    next();
});

app.get('/jwt', (_, res) => {
    res.statusCode = 200;
    res.send(getJWTToken());
});

app.listen(port, () => {
    console.log(`Listening on port ${port}`)
});

const getJWTToken = () => {
    const privateKeyFile = fs.readFileSync(privateKeyFilePath);

    const privateKeyObject = crypto.createPrivateKey({ key: privateKeyFile, format: 'pem' });
    const privateKey = privateKeyObject.export({ format: 'pem', type: 'pkcs8' });

    const publicKeyObject = crypto.createPublicKey({ key: privateKey, format: 'pem' });
    const publicKey = publicKeyObject.export({ format: 'der', type: 'spki' });
    const publicKeyFingerprint = 'SHA256:' + crypto.createHash('sha256') .update(publicKey, 'utf8') .digest('base64');

    const signOptions = {
        iss : qualifiedUsername+ '.' + publicKeyFingerprint,
        sub: qualifiedUsername,
        exp: Math.floor(Date.now() / 1000) + (60 * 60),
    };

    return jwt.sign(signOptions,  privateKey, {algorithm:'RS256'});
};

