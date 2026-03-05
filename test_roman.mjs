import fs from 'fs';

const tests = [
    "CHENNAI INSTITUTE OF TECHNOLOGY Internal Assesment II CO 4 CO 5",
    "CHENNAI INSTITUTE Internal Assessinent - 2",
    "Internal Examination I I  Question",
    "IA - II",
    "IA  2"
];

function detectIAType(text) {
    if (/Internal.{1,30}?(?:3|III|I\s*I\s*I)\b/i.test(text) ||
        /\bIA\s*[-–—]?\s*(?:3|III|I\s*I\s*I)\b/i.test(text) ||
        /Model\s+Exam/i.test(text)) {
        return 'IA3';
    }

    if (/Internal.{1,30}?(?:2|II|I\s*I)\b(?!\s*I)/i.test(text) ||
        /\bIA\s*[-–—]?\s*(?:2|II|I\s*I)\b(?!\s*I)/i.test(text)) {
        return 'IA2';
    }

    if (/Internal.{1,30}?(?:1|I)\b(?!\s*(?:[IV]|I))/i.test(text) ||
        /\bIA\s*[-–—]?\s*(?:1|I)\b(?!\s*(?:[IV]|I))/i.test(text) ||
        /\bEPC\b/i.test(text)) {
        return 'IA1';
    }

    return 'none';
}

for (const t of tests) {
    console.log(`"${t}" => ${detectIAType(t)}`);
}
