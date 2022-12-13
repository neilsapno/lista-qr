import { useState, useEffect } from 'react';
import './App.css';

//MSAL
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { loginRequest } from "./Config";
import { callMsGraph } from "./graph";

//assets
import iredoc from './assets/img/iredoc.png'
import Button from "react-bootstrap/Button";

//components
import QRCoode from './QrCode'

function App() {
  let [name, setName] = useState("")
  let [midName, setMidName] = useState("")
  let [studentNum, setStudent] = useState("")
  let [guild, setGuild] = useState("")
  let [section, setSection] = useState("")
  let [data, setData] = useState({});
  let [isQRShown, setIsQRShown] = useState(false);
  let [isLoggedIn, setIsLoggedIn] = useState(false);
  let [alreadySubmitted, setAlreadySubmitted] = useState(false);
  let [submitButtonText, setSubmitButtonText] = useState("Submit");

  let [inputErrors, setInputErrors] = useState({
    midName: '',
    studentNum: '',
    guild: '',
    section: ''
  });

  const inputHandler = (e) => {
    if (e.target.id === 'midName') {
      setMidName(e.target.value)
    }
    else if (e.target.id === 'studentNum') {
      const value = e.target.value.replace(/\D/g, '');
      setStudent(value)
    }
    else if (e.target.id === 'guild') {
      setGuild(e.target.value)
    }
    else if (e.target.id === 'section') {
      setSection(e.target.value)
    }
    setInputErrors(prevV => {
      return { ...prevV, [e.target.id]: '' }
    })
  }

  const submitData = (e) => {
    e.preventDefault();
    if (isEmptyOrSpaces(studentNum) || isEmptyOrSpaces(guild) || isEmptyOrSpaces(section)) {
      if (isEmptyOrSpaces(studentNum)) {
        setInputErrors({
          ...inputErrors,
          studentNum: "This field cannot be empty"
        });
      }
      else if (isEmptyOrSpaces(guild)) {
        setInputErrors({
          ...inputErrors,
          guild: "This field cannot be empty"
        });
      }
      else if (isEmptyOrSpaces(section)) {
        setInputErrors({
          ...inputErrors,
          section: "This field cannot be empty"
        });
      }
    }
    else {
      setData({
        name: name,
        midName: midName,
        studentNum: studentNum,
        guild: guild,
        section: section
      })
      setIsQRShown(true);

      setAlreadySubmitted(true);
      setSubmitButtonText("Already submitted");
    }
  }

  const { instance, accounts } = useMsal();
  const [graphData, setGraphData] = useState(null);
  const isAuthenticated = useIsAuthenticated();
  const [accessToken, setAccessToken] = useState(null);


  function RequestProfileData() {
    const request = {
      ...loginRequest,
      account: accounts[0]
    };
    instance.acquireTokenSilent(request).then((response) => {
      setAccessToken(response.accessToken);
      callMsGraph(response.accessToken).then(response => setGraphData(response));
    }).catch((e) => {
      instance.acquireTokenRedirect(request).then((response) => {
        setAccessToken(response.accessToken);
        callMsGraph(response.accessToken).then(response => setGraphData(response));
      });
    });
  }

  const handleLogin = (loginType) => {
    if (loginType === "terms") {
      document.querySelector('.terms-box').style.display = "block";
    }
    else if (loginType === "redirect") {
      instance.loginRedirect(loginRequest).catch(e => {
        //console.log(e);
      });

    }
  }

  const handleLogout = (logoutType) => {
    if (logoutType === "redirect") {
      instance.logoutRedirect({
        postLogoutRedirectUri: "/lista-qr",
      });
    }
  }

  useEffect(() => {
    if (isAuthenticated) {
      setIsLoggedIn(true);
      if (isLoggedIn && name.length < 2) {
        RequestProfileData();
      }
    }
    else {
      setIsLoggedIn(false);
    }

    if (accessToken) {
      setName(graphData.surname + ", " + graphData.givenName)
    }
    // eslint-disable-next-line
  }, [graphData, isAuthenticated, isLoggedIn, name.length]);


  return (
    <div className="App">
      <div>
        <hgroup>
          <h1>LISTA QR Code Registration</h1>
          <h3>By Pinoy Big Programmers of STI Sta. Maria</h3>
        </hgroup>
        {accessToken ?
          <div div >
            <form action='#' onSubmit={submitData}>
              <p id='notes'>(*) means required fields.</p>
              <div className="group">
                <input id='name' type="text" value={name} disabled required /><span className="highlight" /><span className="bar" />
                <label>&thinsp; Name*</label>
              </div>
              <div className="group">
                <input onChange={inputHandler} id='midName' type="text" value={midName} required /><span className="highlight" /><span className="bar" />
                <label>Middle Name</label><span className='error'>{inputErrors.midName} </span>
              </div>
              <div className="group">
                <input onChange={inputHandler} id='studentNum' type="text" value={studentNum} required /><span className="highlight" /><span className="bar" />
                <label>Student Number*</label><span className='error'>{inputErrors.studentNum} </span>
              </div>
              <div className="group">
                <select onChange={inputHandler} value={guild} id="guild" required="required">
                  <option id='firstOpt' value="" disabled="disabled"></option>
                  <optgroup label="LISTO"></optgroup>
                  <optgroup></optgroup>
                  <option value="ETIKA">ETIKA</option>
                  <optgroup></optgroup>
                  <option value="IREDOC">IREDOC</option>
                  <optgroup></optgroup>
                  <option value="LETRA">LETRA</option>
                  <optgroup></optgroup>
                  <option value="NUMERIKA">NUMERIKA</option>
                  <optgroup></optgroup>
                  <option value="SWES">SWES</option>
                  <optgroup></optgroup>
                  <optgroup label="GILAS"></optgroup>
                  <optgroup></optgroup>
                  <option value="AWIT">AWIT</option>
                  <optgroup></optgroup>
                  <option value="SINING (MULTIMEDIA)">SINING (MULTIMEDIA)</option>
                  <optgroup></optgroup>
                  <option value="SINING (VISUAL ARTS)">SINING (VISUAL ARTS)</option>
                  <optgroup></optgroup>
                  <option value="GALAW">GALAW</option>
                  <optgroup></optgroup>
                  <option value="INSTRUMENTO">INSTRUMENTO</option>
                  <optgroup></optgroup>
                  <option value="LITERATURA">LITERATURA</option>
                  <optgroup></optgroup>
                </select>
                <label>&thinsp; Guild*</label><span className='error'>{inputErrors.guild} </span>
              </div>
              <div className="group">
                <select onChange={inputHandler} value={section} id="section" required="required">
                  <option id='firstOpt' value="" disabled="disabled"></option>
                  <optgroup label="Grade 11"></optgroup>
                  <optgroup></optgroup>
                  <option value="ABM1101">ABM 1101</option>
                  <optgroup></optgroup>
                  <option value="HUMMS1101">HUMMS 1101</option>
                  <optgroup></optgroup>
                  <option value="STEM1101">STEM 1101</option>
                  <optgroup></optgroup>
                  <option value="STEM1102">STEM 1102</option>
                  <optgroup></optgroup>
                  <option value="ITM1101">ITM 1101</option>
                  <optgroup></optgroup>
                  <option value="GAS1101">GAS 1101</option>
                  <optgroup></optgroup>
                  <option value="DA1101">DA 1101</option>
                  <optgroup></optgroup>
                  <option value="CA1101">CA 1101</option>
                  <optgroup></optgroup>
                  <option value="TO1101">TO 1101</option>
                  <optgroup></optgroup>
                  <optgroup label='Grade 12'></optgroup>
                  <optgroup></optgroup>
                  <option value="ABM1201">ABM 1201</option>
                  <optgroup></optgroup>
                  <option value="HUMMS1201">HUMMS 1201</option>
                  <optgroup></optgroup>
                  <option value="STEM1201">STEM 1201</option>
                  <optgroup></optgroup>
                  <option value="STEM1202">STEM 1202</option>
                  <optgroup></optgroup>
                  <option value="ITM1201">ITM 1201</option>
                  <optgroup></optgroup>
                  <option value="GAS1201">GAS 1201</option>
                  <optgroup></optgroup>
                  <option value="DA1201">DA 1201</option>
                  <optgroup></optgroup>
                  <option value="CA1201">CA 1201</option>
                  <optgroup></optgroup>
                  <option value="TO1201">TO 1201</option>
                </select>
                <label>&thinsp; Section*</label><span className='error'>{inputErrors.section} </span>
              </div>
              <button onClick={submitData} disabled={alreadySubmitted} type="submit" className="button buttonBlue">{submitButtonText}
                <div className="ripples buttonRipples"><span className="ripplesCircle" /></div>
              </button>
            </form>

            <Button variant="primary" style={styles.signButton} className="signButton" onClick={() => { handleLogout("redirect") }}>SIGN OUT</Button>
          </div>
          :
          <div>
            <div className="terms-box">
              <div className="terms-text">
                <h2 style={{ textAlign: "center" }}><b>PRIVACY POLICY</b></h2>
                <p>Effective date: 2022 - 2023</p>
                <p>1. <b>Introduction</b></p>
                <p>Welcome to <b> LISTA</b> by <b>PBP</b>.</p>
                <p><b>PBP</b> (“us”, “we”, or “our”) operates <b>https://pinoybigprogrammers.github.io/lista-qr/</b> (hereinafter referred to as <b>“Service”</b>).</p>
                <p>Our Privacy Policy governs your visit to <b>https://pinoybigprogrammers.github.io/lista-qr/</b>, and explains how we collect, safeguard and disclose information that results from your use of our Service.</p>
                <p>We use your data to provide and improve Service. By using Service, you agree to the collection and use of information in accordance with this policy. Unless otherwise defined in this Privacy Policy, the terms used in this Privacy Policy have the same meanings as in our Terms and Conditions.</p>
                <p>Our Terms and Conditions (<b>“Terms”</b>) govern all use of our Service and together with the Privacy Policy constitutes your agreement with us (<b>“agreement”</b>).</p>
                <p>2. <b>Definitions</b></p>
                <p><b>SERVICE</b> means the https://pinoybigprogrammers.github.io/lista-qr/ website operated by PBP.</p>
                <p><b>PERSONAL DATA</b> means data about a living individual who can be identified from those data (or from those and other information either in our possession or likely to come into our possession).</p>
                <p><b>USAGE DATA</b> is data collected automatically either generated by the use of Service or from Service infrastructure itself (for example, the duration of a page visit).</p>
                <p><b>COOKIES</b> are small files stored on your device (computer or mobile device).</p>
                <p><b>DATA CONTROLLER</b> means a natural or legal person who (either alone or jointly or in common with other persons) determines the purposes for which and the manner in which any personal data are, or are to be, processed. For the purpose of this Privacy Policy, we are a Data Controller of your data.</p>
                <p><b>DATA PROCESSORS (OR SERVICE PROVIDERS)</b> means any natural or legal person who processes the data on behalf of the Data Controller. We may use the services of various Service Providers in order to process your data more effectively.</p> <p><b>DATA SUBJECT</b> is any living individual who is the subject of Personal Data.</p>
                <p><b>THE USER</b> is the individual using our Service. The User corresponds to the Data Subject, who is the subject of Personal Data.</p>
                <p>3. <b>Information Collection and Use</b></p>
                <p>We collect several different types of information for various purposes to provide and improve our Service to you.</p>
                <p>4. <b>Types of Data Collected</b></p>
                <p><b>Personal Data</b></p>
                <p>While using our Service, we may ask you to provide us with certain personally identifiable information that can be used to contact or identify you (<b>“Personal Data”</b>). Personally identifiable information may include, but is not limited to:</p>
                <p>0.1. Email address</p>
                <p>0.2. First name and last name</p>
                <p>0.3. Phone number</p>
                <p>0.4. Cookies and Usage Data</p>
                <p>We may use your Personal Data to contact you with newsletters, marketing or promotional materials and other information that may be of interest to you. You may opt out of receiving any, or all, of these communications from us by following the unsubscribe link.</p>
                <p><b>Usage Data</b></p>
                <p>We may also collect information that your browser sends whenever you visit our Service or when you access Service by or through any device (<b>“Usage Data”</b>).</p>
                <p>This Usage Data may include information such as your computer’s Internet Protocol address (e.g. IP address), browser type, browser version, the pages of our Service that you visit, the time and date of your visit, the time spent on those pages, unique device identifiers and other diagnostic data.</p>
                <p>When you access Service with a device, this Usage Data may include information such as the type of device you use, your device unique ID, the IP address of your device, your device operating system, the type of Internet browser you use, unique device identifiers and other diagnostic data.</p>

                <p><b>Tracking Cookies Data</b></p>
                <p>We use cookies and similar tracking technologies to track the activity on our Service and we hold certain information.</p>
                <p>Cookies are files with a small amount of data which may include an anonymous unique identifier. Cookies are sent to your browser from a website and stored on your device. Other tracking technologies are also used such as beacons, tags and scripts to collect and track information and to improve and analyze our Service.</p>
                <p>You can instruct your browser to refuse all cookies or to indicate when a cookie is being sent. However, if you do not accept cookies, you may not be able to use some portions of our Service.</p>
                <p>Examples of Cookies we use:</p>
                <p>0.1. <b>Session Cookies:</b> We use Session Cookies to operate our Service.</p>
                <p>0.2. <b>Preference Cookies:</b> We use Preference Cookies to remember your preferences and various settings.</p>
                <p>0.3. <b>Security Cookies:</b> We use Security Cookies for security purposes.</p>
                <p>0.4. <b>Advertising Cookies:</b> Advertising Cookies are used to serve you with advertisements that may be relevant to you and your interests.</p>
                <p><b>Other Data</b></p>
                <p>While using our Service, we may also collect the following information: sex, age, date of birth, place of birth, passport details, citizenship, registration at place of residence and actual address, telephone number (work, mobile), details of documents on education, qualification, professional training, employment agreements, <a href="https://policymaker.io/non-disclosure-agreement/">NDA agreements</a>, information on bonuses and compensation, information on marital status, family members, social security (or other taxpayer identification) number, office location and other data.</p>
                <p>5. <b>Use of Data</b></p>
                <p>LISTA uses the collected data for various purposes:</p>
                <p>0.1. to provide and maintain our Service;</p>
                <p>0.2. to notify you about changes to our Service;</p>
                <p>0.3. to allow you to participate in interactive features of our Service when you choose to do so;</p>
                <p>0.4. to provide customer support;</p>
                <p>0.5. to gather analysis or valuable information so that we can improve our Service;</p>
                <p>0.6. to monitor the usage of our Service;</p>
                <p>0.7. to detect, prevent and address technical issues;</p>
                <p>0.8. to fulfil any other purpose for which you provide it;</p>
                <p>0.9. to carry out our obligations and enforce our rights arising from any contracts entered into between you and us, including for billing and collection;</p>
                <p>0.10. to provide you with notices about your account and/or subscription, including expiration and renewal notices, email-instructions, etc.;</p>
                <p>0.11. to provide you with news, special offers and general information about other goods, services and events which we offer that are similar to those that you have already purchased or enquired about unless you have opted not to receive such information;</p>
                <p>0.12. in any other way we may describe when you provide the information;</p>
                <p>0.13. for any other purpose with your consent.</p>
                <p>6. <b>Retention of Data</b></p>
                <p>We will retain your Personal Data only for as long as is necessary for the purposes set out in this Privacy Policy. We will retain and use your Personal Data to the extent necessary to comply with our legal obligations (for example, if we are required to retain your data to comply with applicable laws), resolve disputes, and enforce our legal agreements and policies.</p>
                <p>We will also retain Usage Data for internal analysis purposes. Usage Data is generally retained for a shorter period, except when this data is used to strengthen the security or to improve the functionality of our Service, or we are legally obligated to retain this data for longer time periods.</p>
                <p>7. <b>Transfer of Data</b></p>
                <p>Your information, including Personal Data, may be transferred to – and maintained on – computers located outside of your state, province, country or other governmental jurisdiction where the data protection laws may differ from those of your jurisdiction.</p>
                <p>If you are located outside Philippines and choose to provide information to us, please note that we transfer the data, including Personal Data, to Philippines and process it there.</p>
                <p>Your consent to this Privacy Policy followed by your submission of such information represents your agreement to that transfer.</p>
                <p>PBP will take all the steps reasonably necessary to ensure that your data is treated securely and in accordance with this Privacy Policy and no transfer of your Personal Data will take place to an organisation or a country unless there are adequate controls in place including the security of your data and other personal information.</p>
                <p>8. <b>Disclosure of Data</b></p>
                <p>We may disclose personal information that we collect, or you provide:</p>
                <p>0.1. <b>Business Transaction.</b></p><p>If we or our subsidiaries are involved in a merger, acquisition or asset sale, your Personal Data may be transferred.</p><p>0.2. <b>Other cases. We may disclose your information also:</b></p><p>0.2.1. to our subsidiaries and affiliates;</p><p>0.2.2. to contractors, service providers, and other third parties we use to support our business;</p><p>0.2.3. to fulfill the purpose for which you provide it;</p><p>0.2.4. for the purpose of including your company’s logo on our website;</p><p>0.2.5. for any other purpose disclosed by us when you provide the information;</p><p>0.2.6. with your consent in any other cases;</p><p>0.2.7. if we believe disclosure is necessary or appropriate to protect the rights, property, or safety of the Company, our customers, or others.</p>
                <p>9. <b>Security of Data</b></p>
                <p>The security of your data is important to us but remember that no method of transmission over the Internet or method of electronic storage is 100% secure. While we strive to use commercially acceptable means to protect your Personal Data, we cannot guarantee its absolute security.</p>
                <p>10. <b>Your Data Protection Rights Under General Data Protection Regulation (GDPR)
                </b></p>
                <p>If you are a resident of the European Union (EU) and European Economic Area (EEA), you have certain data protection rights, covered by GDPR.</p>
                <p>We aim to take reasonable steps to allow you to correct, amend, delete, or limit the use of your Personal Data.</p>
                <p> If you wish to be informed what Personal Data we hold about you and if you want it to be removed from our systems, please email us at <b>sapno.244846@stamaria.sti.edu.ph</b>.</p>
                <p>In certain circumstances, you have the following data protection rights:</p>
                <p>0.1. the right to access, update or to delete the information we have on you;</p>
                <p>0.2. the right of rectification. You have the right to have your information rectified if that information is inaccurate or incomplete;</p>
                <p>0.3. the right to object. You have the right to object to our processing of your Personal Data;</p>
                <p>0.4. the right of restriction. You have the right to request that we restrict the processing of your personal information;</p>
                <p>0.5. the right to data portability. You have the right to be provided with a copy of your Personal Data in a structured, machine-readable and commonly used format;</p>
                <p>0.6. the right to withdraw consent. You also have the right to withdraw your consent at any time where we rely on your consent to process your personal information;</p>
                <p>Please note that we may ask you to verify your identity before responding to such requests. Please note, we may not able to provide Service without some necessary data.</p>
                <p>You have the right to complain to a Data Protection Authority about our collection and use of your Personal Data. For more information, please contact your local data protection authority in the European Economic Area (EEA).</p>
                <p>11. <b>Your Data Protection Rights under the California Privacy Protection Act (CalOPPA)</b></p>
                <p>CalOPPA is the first state law in the nation to require commercial websites and online services to post a privacy policy. The law’s reach stretches well beyond California to require a person or company in the United States (and conceivable the world) that operates websites collecting personally identifiable information from California consumers to post a conspicuous privacy policy on its website stating exactly the information being collected and those individuals with whom it is being shared, and to comply with this policy.</p>
                <p>According to CalOPPA we agree to the following:</p>
                <p>0.1. users can visit our site anonymously;</p>
                <p>0.2. our Privacy Policy link includes the word “Privacy”, and can easily be found on the home page of our website;</p>
                <p>0.3. users will be notified of any privacy policy changes on our Privacy Policy Page;</p>
                <p>0.4. users are able to change their personal information by emailing us at <b>sapno.244846@stamaria.sti.edu.ph</b>.</p>
                <p>Our Policy on “Do Not Track” Signals:</p>
                <p>We honor Do Not Track signals and do not track, plant cookies, or use advertising when a Do Not Track browser mechanism is in place. Do Not Track is a preference you can set in your web browser to inform websites that you do not want to be tracked.</p>
                <p>You can enable or disable Do Not Track by visiting the Preferences or Settings page of your web browser.</p>
                <p>12. <b>Your Data Protection Rights under the California Consumer Privacy Act (CCPA)</b></p>
                <p>If you are a California resident, you are entitled to learn what data we collect about you, ask to delete your data and not to sell (share) it. To exercise your data protection rights, you can make certain requests and ask us:</p>
                <p><b>0.1. What personal information we have about you. If you make this request, we will return to you:</b></p>
                <p>0.0.1. The categories of personal information we have collected about you.</p>
                <p>0.0.2. The categories of sources from which we collect your personal information.</p>
                <p>0.0.3. The business or commercial purpose for collecting or selling your personal information.</p>
                <p>0.0.4. The categories of third parties with whom we share personal information.</p>
                <p>0.0.5. The specific pieces of personal information we have collected about you.</p>
                <p>0.0.6. A list of categories of personal information that we have sold, along with the category of any other company we sold it to. If we have not sold your personal information, we will inform you of that fact.</p>
                <p>0.0.7. A list of categories of personal information that we have disclosed for a business purpose, along with the category of any other company we shared it with.</p>
                <p>Please note, you are entitled to ask us to provide you with this information up to two times in a rolling twelve-month period. When you make this request, the information provided may be limited to the personal information we collected about you in the previous 12 months.</p>
                <p><b>0.2. To delete your personal information. If you make this request, we will delete the personal information we hold about you as of the date of your request from our records and direct any service providers to do the same. In some cases, deletion may be accomplished through de-identification of the information. If you choose to delete your personal information, you may not be able to use certain functions that require your personal information to operate.</b></p>
                <p><b>0.3. To stop selling your personal information. We don’t sell or rent your personal information to any third parties for any purpose. We do not sell your personal information for monetary consideration. However, under some circumstances, a transfer of personal information to a third party, or within our family of companies, without monetary consideration may be considered a “sale” under California law. You are the only owner of your Personal Data and can request disclosure or deletion at any time.</b></p>
                <p>If you submit a request to stop selling your personal information, we will stop making such transfers.</p>
                <p>Please note, if you ask us to delete or stop selling your data, it may impact your experience with us, and you may not be able to participate in certain programs or membership services which require the usage of your personal information to function. But in no circumstances, we will discriminate against you for exercising your rights.</p>
                <p>To exercise your California data protection rights described above, please send your request(s) by email: <b>sapno.244846@stamaria.sti.edu.ph</b>.</p>
                <p>Your data protection rights, described above, are covered by the CCPA, short for the California Consumer Privacy Act. To find out more, visit the official California Legislative Information website. The CCPA took effect on 01/01/2020.</p>
                <p>13. <b>Service Providers</b></p>
                <p>We may employ third party companies and individuals to facilitate our Service (<b>“Service Providers”</b>), provide Service on our behalf, perform Service-related services or assist us in analysing how our Service is used.</p>
                <p>These third parties have access to your Personal Data only to perform these tasks on our behalf and are obligated not to disclose or use it for any other purpose.</p>
                <p>14. <b>Analytics</b></p>
                <p>We may use third-party Service Providers to monitor and analyze the use of our Service.</p>
                <p>15. <b>CI/CD tools</b></p>
                <p>We may use third-party Service Providers to automate the development process of our Service.</p>

                <p>16. <b>Behavioral Remarketing</b></p>
                <p>We may use remarketing services to advertise on third party websites to you after you visited our Service. We and our third-party vendors use cookies to inform, optimise and serve ads based on your past visits to our Service.</p>

                <p>17. <b>Links to Other Sites</b></p>
                <p>Our Service may contain links to other sites that are not operated by us. If you click a third party link, you will be directed to that third party’s site. We strongly advise you to review the Privacy Policy of every site you visit.</p>
                <p>We have no control over and assume no responsibility for the content, privacy policies or practices of any third party sites or services.</p>
                <p>18. <b><b>Children’s Privacy</b></b></p>
                <p>Our Services are not intended for use by children under the age of 18 (<b>“Child”</b> or <b>“Children”</b>).</p>
                <p>We do not knowingly collect personally identifiable information from Children under 18. If you become aware that a Child has provided us with Personal Data, please contact us. If we become aware that we have collected Personal Data from Children without verification of parental consent, we take steps to remove that information from our servers.</p>
                <p>19. <b>Changes to This Privacy Policy</b></p>
                <p>We may update our Privacy Policy from time to time. We will notify you of any changes by posting the new Privacy Policy on this page.</p>
                <p>We will let you know via email and/or a prominent notice on our Service, prior to the change becoming effective and update “effective date” at the top of this Privacy Policy.</p>
                <p>You are advised to review this Privacy Policy periodically for any changes. Changes to this Privacy Policy are effective when they are posted on this page.</p>
                <p>20. <b>Contact Us</b></p>
                <p>If you have any questions about this Privacy Policy, please contact us by email: <b>sapno.244846@stamaria.sti.edu.ph</b>.</p>
              </div>
              <h4>I Agree to the <span>Terms of Service</span> and I read the Privacy Notice.</h4>
              <div className="buttons">
                <button className="btnt red-btnt" onClick={() => handleLogin("redirect")}>Accept</button>
                <button className="btnt gray-btnt" onClick={() => document.querySelector('.terms-box').style.display = "none"}>Decline</button>
              </div>
            </div>
            <Button variant="primary" style={styles.signButton} className="signButton btnt" onClick={() => handleLogin("terms")}>SIGN IN WITH YOUR STI 0365 ACCOUNT</Button>
          </div>
        }
        {isQRShown && <QRCoode data={data} />
        }
        <footer><a href="https://www.facebook.com/profile.php?id=100087772660179" rel="noopener noreferrer" target="_blank"><img alt="IREDOC#1" src={iredoc} /></a>
          <p>Pinoy Big Programmers (PBP). All Rights Reserved.</p>
        </footer>
      </div>
    </div>
  );
}
const styles = {

  signButton: {
    background: '#eb4626',
  }

}

function isEmptyOrSpaces(str) {
  return str === null || str.match(/^ *$/) !== null;
}

/*Pre ung terms and service baka malimot*/

export default App;