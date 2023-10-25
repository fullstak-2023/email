#!/usr/bin/env node

const ShanaTova = require('../7.email.js');

// יבוא השם שהקלדתי למטה והכנסתו למשתנה
const userName = process.argv[2]
const userEmail = process.argv[3]

ShanaTova.getNameAndEmail(userName, userEmail);
