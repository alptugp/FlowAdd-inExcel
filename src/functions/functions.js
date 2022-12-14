import fetch from "node-fetch";
window.sharedState = new Array();
/* eslint-disable @typescript-eslint/no-unused-vars */
/* global console setInterval, clearInterval */

/**
 * Pulls the value of the corresponding parameter from Flow.
 * @customfunction
 * @param {string} parameterName
 * @returns The value of the given parameter.
 */
async function flow(parameterName) {
  let username = window.sharedState[0];
  let password = window.sharedState[1];
  let projectName = window.sharedState[2];

  const query1 = {
    ClientId: "3asjpt4hmudvll6us1v45i1vs3",
    AuthFlow: "USER_PASSWORD_AUTH",
    AuthParameters: {
      USERNAME: username,
      PASSWORD: password,
    },
  };

  let body = JSON.stringify(query1);

  const cognitoUrl = "https://cognito-idp.eu-west-2.amazonaws.com/eu-west-2_XJ0fMS4Ey";

  const request = fetch(cognitoUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-amz-json-1.1",
      "X-Amz-Target": "AWSCognitoIdentityProviderService.InitiateAuth",
    },
    body: body,
  });

  let idToken;

  idToken = await request.then((res) => res.json()).then((data) => data["AuthenticationResult"]["IdToken"]);

  const projectQuery = `
    query Projects {
      project {
      project_id
      name
      description
      creator {
        user_id
        given_name
        family_name
      }
      archived
    }
    }
    `;

  let projectBody = JSON.stringify({ query: projectQuery });
  const queryUrl = "https://staging.api.flowengineering.com/v1/graphql";

  var bearer = "Bearer " + idToken;

  const projectQueryResult = fetch(queryUrl, {
    method: "POST",
    headers: {
      Authorization: bearer,
      "Content-Type": "application/json",
    },
    body: projectBody,
  });

  let projects = await projectQueryResult.then((res) => res.json()).then((data) => data["data"]["project"]);

  let projectId = getMatchingProjectId(projects, projectName);

  const categoryQuery = `
query DataCategories($projectId: uuid!) {
  data_category(where: { project_id: { _eq: $projectId } }) {
    category_id
    name
    human_id_prefix
    project {
      project_id
      name
    }
    archived
  }
}
`;

  let categoryBody = JSON.stringify({ query: categoryQuery, variables: { projectId } });

  const categoryQueryResult = fetch(queryUrl, {
    method: "POST",
    headers: {
      Authorization: bearer,
      "Content-Type": "application/json",
    },
    body: categoryBody,
  });

  let categoryId = await categoryQueryResult
    .then((res) => res.json())
    .then((data) => data["data"]["data_category"][0]["category_id"]);

  const dataQuery = `
query Data($categoryId: uuid!) {
  data(where: { category_id: { _eq: $categoryId } }) {
    data_id
    name
    human_id
    value
    category {
      category_id
      name
      human_id_prefix
    }
    archived
  }
}
`;

  let dataBody = JSON.stringify({
    query: dataQuery,
    variables: { categoryId },
  });

  const dataQueryResult = fetch(queryUrl, {
    method: "POST",
    headers: {
      Authorization: bearer,
      "Content-Type": "application/json",
    },
    body: dataBody,
  });

  let datas = await dataQueryResult.then((res) => res.json()).then((data) => data["data"]["data"]);

  let parameterVal = getMatchingDataVal(datas, parameterName);

  return parameterVal;
}

function getMatchingProjectId(projects, projectName) {
  let projectId;
  for (let project of projects) {
    if (project["name"] == projectName) {
      projectId = project["project_id"];
    }
  }
  if (projectId == null) {
    console.log("Please enter a project name which exists in Flow.");
    return;
  }
  return projectId;
}

function getMatchingDataVal(datas, parameterName) {
  let parameterValue;
  for (let data of datas) {
    if (data["name"] == parameterName) {
      parameterValue = data["value"];
    }
  }
  if (parameterValue == null) {
    console.log("Please enter a parameter name which exists in Flow.");
    return;
  }
  return parameterValue;
}
