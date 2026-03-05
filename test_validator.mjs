import fs from 'fs';

const htmlContent = `
<div class="pdf-content">
    <table style="width: 100%; border-collapse: collapse; border: 1px solid black; margin-bottom: 20px;">
      <tr>
        <td colspan="4" style="text-align: center; border: 1px solid black; padding: 10px;">
          <h2 style="margin: 0;">CHENNAI INSTITUTE OF TECHNOLOGY</h2>
          <p style="margin: 5px 0;">Autonomous</p>
          <p style="margin: 2px 0;">Sarathy Nagar, Kundrathur, Chennai - 600 069.</p>
          <p style="margin: 5px 0;" id="assTypeCell"><strong>Internal Assessment 2</strong></p>
        </td>
      </tr>
      <tr>
        <td style="border: 1px solid black; padding: 5px;">Date: 13-02-2026</td>
        <td style="border: 1px solid black; padding: 5px;" id="code">Subject Code / Name: MT3602 / ROBOTICS AND MACHINE VISION SYSTEMS</td>
        <td style="border: 1px solid black; padding: 5px;">Max. Marks: 50 Marks</td>
        <td style="border: 1px solid black; padding: 5px;">Time: 1.30 hrs</td>
      </tr>
      <tr>
        <td style="border: 1px solid black; padding: 5px;" colspan="2" id="branchCell">Branch: Mechatronics Engineering</td>
        <td style="border: 1px solid black; padding: 5px;" colspan="2" id="yearSemesterCell">Year / Sem: III / VI</td>
      </tr>
    </table>
    <div id="courseObjectives" style="display:none;">["Understand the principles of robotics","Gain proficiency","Analyse the complex robotics","Apply the theoretical knowledge","Keep up-to-date with"]</div>
    <div id="courseOutcomes" style="display:none;">[{"co":"CO1","description":"Express the basic concepts","level":"L2"},{"co":"CO2","description":"Explain the types","level":"L2"},{"co":"CO3","description":"Evaluate the kinematic","level":"L5"}]</div>
</div>
Some other text that has Course Objectives and CO 5 and Internal Assesment I I.
`;

function testLogic() {
    const text = htmlContent
        .replace(/<[^>]+>/g, ' ')
        .replace(/&nbsp;/gi, ' ')
        .replace(/&amp;/gi, '&')
        .replace(/\s+/g, ' ');

    console.log("Extracted text: ", text.substring(0, 200));

    let ia3Match = /Internal.{1,30}?(?:3|III|I\s*I\s*I|l\s*l\s*l)\b/i.exec(text);
    console.log("IA3 Match from Regex:", ia3Match ? ia3Match[0] : null);

    let ia2Match = /Internal.{1,30}?(?:2|II|I\s*I|l\s*l)\b(?!\s*(?:I|l))/i.exec(text);
    console.log("IA2 Match from Regex:", ia2Match ? ia2Match[0] : null);
}

testLogic();
