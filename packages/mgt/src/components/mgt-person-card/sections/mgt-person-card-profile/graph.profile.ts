/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { PhysicalAddress } from '@microsoft/microsoft-graph-types-beta';
import { IGraph } from '@microsoft/mgt-element';

/**
 * Represents the details of meaningful dates associated with a person in a user's profile.
 *
 * @interface IPersonAnniversary
 */
export interface IPersonAnniversary {
  /**
   * Contains the date associated with the anniversary type.
   *
   * @type {Date}
   * @memberof IPersonAnniversary
   */
  date: Date;
  /**
   * Possible values are: birthday, wedding, unknownFutureValue.
   *
   * @type {string}
   * @memberof IPersonAnniversary
   */
  type: string;
}

/**
 * Provides detailed information about interests the user has associated with themselves in various services.
 *
 * @interface IPersonInterest
 */
export interface IPersonInterest {
  /**
   * Contains categories a user has associated with the interest (for example, personal, recipies).
   *
   * @type {string[]}
   * @memberof IPersonInterest
   */
  categories: string[];
  /**
   * Contains a description of the interest.
   *
   * @type {string}
   * @memberof IPersonInterest
   */
  description: string;
  /**
   * Contains a friendly name for the interest.
   *
   * @type {string}
   * @memberof IPersonInterest
   */
  displayName: string;
  /**
   * Contains a link to a web page or resource about the interest.
   *
   * @type {string}
   * @memberof IPersonInterest
   */
  webUrl: string;
}

/**
 * Represents detailed information about languages that a user has added to their profile.
 *
 * @interface ILanguageProficiency
 */
export interface ILanguageProficiency {
  /**
   * Contains the long-form name for the language.
   *
   * @type {string}
   * @memberof ILanguageProficiency
   */
  displayName: string;
  /**
   * Possible values are: elementary, conversational, limitedWorking, professionalWorking,
   * fullProfessional, nativeOrBilingual, unknownFutureValue.
   *
   * @type {string}
   * @memberof ILanguageProficiency
   */
  proficiency: string;
  /**
   * Contains the four-character BCP47 name for the language (en-US, no-NB, en-AU).
   *
   * @type {string}
   * @memberof ILanguageProficiency
   */
  tag: string;
}

/**
 * Represents detailed information about skills associated with a user in various services.
 *
 * @interface ISkillProficiency
 */
export interface ISkillProficiency {
  /**
   * Contains categories a user has associated with the skill (for example, personal, professional, hobby).
   *
   * @type {string[]}
   * @memberof ISkillProficiency
   */
  categories: string[];
  /**
   * Contains a friendly name for the skill.
   *
   * @type {string}
   * @memberof ISkillProficiency
   */
  displayName: string;
  /**
   * Possible values are: elementary, limitedWorking, generalProfessional, advancedProfessional, expert, unknownFutureValue.
   *
   * @type {string}
   * @memberof ISkillProficiency
   */
  proficiency: string;
  /**
   * Contains a link to an information source about the skill.
   *
   * @type {string}
   * @memberof ISkillProficiency
   */
  webUrl: string;
}

/**
 * Represents information about companies related to entities within their profile.
 *
 * @export
 * @interface ICompanyDetail
 */
export interface ICompanyDetail {
  /**
   * Address of the company.
   *
   * @type {PhysicalAddress}
   * @memberof ICompanyDetail
   */
  address: PhysicalAddress;
  /**
   * Department Name within a company.
   *
   * @type {string}
   * @memberof ICompanyDetail
   */
  department: string;
  /**
   * Company name.
   *
   * @type {string}
   * @memberof ICompanyDetail
   */
  displayName: string;
  /**
   * Office Location of the person referred to.
   *
   * @type {string}
   * @memberof ICompanyDetail
   */
  officeLocation: string;
  /**
   * Pronunciation guide for the company name.
   *
   * @type {string}
   * @memberof ICompanyDetail
   */
  pronounciation: string;
  /**
   * Link to the company home page.
   *
   * @type {string}
   * @memberof ICompanyDetail
   */
  webUrl: string;
}

/**
 * Represents information about positions related to entities within a user's profile.
 *
 * @interface IPositionDetail
 */
export interface IPositionDetail {
  /**
   * Detail about the company or employer.
   *
   * @type {ICompanyDetail}
   * @memberof IPositionDetail
   */
  company: ICompanyDetail;
  /**
   * Description of the position in question.
   *
   * @type {string}
   * @memberof IPositionDetail
   */
  description: string;
  /**
   * When the position ended.
   *
   * @type {Date}
   * @memberof IPositionDetail
   */
  endMonthYear: Date;
  /**
   * The title held when in that position.
   *
   * @type {string}
   * @memberof IPositionDetail
   */
  jobTitle: string;
  /**
   * The role the position entailed.
   *
   * @type {string}
   * @memberof IPositionDetail
   */
  role: string;
  /**
   * The start month and year of the position.
   *
   * @type {Date}
   * @memberof IPositionDetail
   */
  startMonthYear: Date;
  /**
   * Short summary of the position.
   *
   * @type {string}
   * @memberof IPositionDetail
   */
  summary: string;
}

/**
 * Represents detailed information about work positions associated with a user's profile.
 *
 * @export
 * @interface IWorkPosition
 */
export interface IWorkPosition {
  /**
   * Contains categories a user has associated with the position (for example, digital transformation, people).
   *
   * @type {string[]}
   * @memberof IWorkPosition
   */
  categories: string[];
  /**
   * Contains detail about the user's current and previous employment positions.
   *
   * @type {IPositionDetail}
   * @memberof IWorkPosition
   */
  detail: IPositionDetail;
}

/**
 * Represents additional detail about an undergraduate, graduate, postgraduate degree
 * or other educational activity that a user has undertaken and is used within an educationalActivity resource.
 *
 * @export
 * @interface IInstitutionData
 */
export interface IInstitutionData {
  /**
   * Short description of the institution the user studied at.
   *
   * @type {string}
   * @memberof IInstitutionData
   */
  description: string;
  /**
   * Name of the institution the user studied at.
   *
   * @type {string}
   * @memberof IInstitutionData
   */
  displayName: string;
  /**
   * Address or location of the institute.
   *
   * @type {PhysicalAddress}
   * @memberof IInstitutionData
   */
  location: PhysicalAddress;
  /**
   * Link to the institution or department homepage.
   *
   * @type {string}
   * @memberof IInstitutionData
   */
  webUrl: string;
}

/**
 * Represents additional detail about an undergraduate, graduate, postgraduate degree
 * or other educational activity that a user has undertaken and is used within an educationalActivity resource.
 *
 * @export
 * @interface IEducationalActivityDetail
 */
export interface IEducationalActivityDetail {
  /**
   * Shortened name of the degree or program (example: PhD, MBA)
   *
   * @type {string}
   * @memberof IEducationalActivityDetail
   */
  abbreviation: string;
  /**
   * Extracurricular activities undertaken alongside the program.
   *
   * @type {string}
   * @memberof IEducationalActivityDetail
   */
  activities: string;
  /**
   * Any awards or honors associated with the program.
   *
   * @type {string}
   * @memberof IEducationalActivityDetail
   */
  awards: string;
  /**
   * Short description of the program provided by the user.
   *
   * @type {string}
   * @memberof IEducationalActivityDetail
   */
  description: string;
  /**
   * Long-form name of the program that the user has provided.
   *
   * @type {string}
   * @memberof IEducationalActivityDetail
   */
  displayName: string;
  /**
   * Majors and minors associated with the program. (if applicable)
   *
   * @type {string}
   * @memberof IEducationalActivityDetail
   */
  fieldsOfStufy: string;
  /**
   * The final grade, class, GPA or score.
   *
   * @type {string}
   * @memberof IEducationalActivityDetail
   */
  grade: string;
  /**
   * Additional notes the user has provided.
   *
   * @type {string}
   * @memberof IEducationalActivityDetail
   */
  notes: string;
  /**
   * Link to the degree or program page.
   *
   * @type {string}
   * @memberof IEducationalActivityDetail
   */
  webUrl: string;
}

/**
 * Represents data that a user has supplied related to undergraduate, graduate, postgraduate or other educational activities.
 *
 * @export
 * @interface IEducationalActivity
 */
export interface IEducationalActivity {
  /**
   * The month and year the user graduated or completed the activity.
   *
   * @type {Date}
   * @memberof IEducationalActivity
   */
  completionMonthYear: Date;
  /**
   * The month and year the user completed the educational activity referenced.
   *
   * @type {Date}
   * @memberof IEducationalActivity
   */
  endMonthYear: Date;
  /**
   * Contains details of the institution studied at.
   *
   * @type {IInstitutionData}
   * @memberof IEducationalActivity
   */
  institution: IInstitutionData;
  /**
   * Contains extended information about the program or course.
   *
   * @type {IEducationalActivityDetail}
   * @memberof IEducationalActivity
   */
  program: IEducationalActivityDetail;
  /**
   * The month and year the user commenced the activity referenced.
   *
   * @type {Date}
   * @memberof IEducationalActivity
   */
  startMonthYear: Date;
}

/**
 * Represents descriptive properties of a person in a tenant.
 * These properties are surfaced in shared people experiences across Microsoft 365
 * and third-party services and experiences via Microsoft Graph.
 *
 * @interface IProfile
 */
export interface IProfile {
  // tslint:disable-next-line: completed-docs
  id: string;
  /**
   * Represents the details of meaningful dates associated with a person.
   *
   * @type {IPersonAnniversary[]}
   * @memberof IProfile
   */
  anniversaries: IPersonAnniversary[];
  /**
   * Represents data that a user has supplied related to undergraduate, graduate, postgraduate or other educational activities.
   *
   * @type {IEducationActivity[]}
   * @memberof IProfile
   */
  educationalActivities: IEducationalActivity[];
  /**
   * Provides detailed information about interests the user has associated with themselves in various services.
   *
   * @type {IPersonInterest[]}
   * @memberof IProfile
   */
  interests: IPersonInterest[];
  /**
   * Represents detailed information about languages that a user has added to their profile.
   *
   * @type {ILanguageProficiency[]}
   * @memberof IProfile
   */
  languages: ILanguageProficiency[];
  /**
   * Represents detailed information about work positions associated with a user's profile.
   *
   * @type {IWorkPosition[]}
   * @memberof IProfile
   */
  positions: IWorkPosition[];
  /**
   * Represents detailed information about skills associated with a user in various services.
   *
   * @type {ISkillProficiency[]}
   * @memberof IProfile
   */
  skills: ISkillProficiency[];
}

/**
 * TODO: Figure out the correct graph call
 *
 * @export
 * @param {IGraph} graph
 * @param {string} userId
 * @returns {IProfile}
 */
export async function getProfile(graph: IGraph, userId: string): Promise<IProfile> {
  const profile = await graph.api(`/users/${userId}/profile`).get();
  return profile;
}
