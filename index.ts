import * as request from 'request-promise';

const fs = require('fs-extra');


import * as _ from 'lodash';
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types"


import {diff} from 'deep-diff';
import {StatusCodeError} from "request-promise/errors";

const Bluebird = require('bluebird');
const getToken = require('./gettoken.js');
const getJobs = require('./getJobs.js');

const SlackWebhook = require('slack-webhook');

const notifier = require('node-notifier');


const dotenv = require('dotenv');

dotenv.config();


// fill in with your things
let token = process.env.TOKEN || '';
const teamsurl = process.env.TEAMSURL;
const slackurl = process.env.SLACKURL;
const file = process.env.DATAFILE;

const slack = new SlackWebhook(slackurl);



setTimeout((function () {
    return process.exit(1);
}), 1000 * 60 * 5);

interface Job {
    title: string;
    division: string;
    location: string;
    id: string;
}

interface Data {
    jobs: Job[];
    users: MicrosoftGraph.User[];
    groups: MicrosoftGraph.Group[];
    usersToGroups: { [i: string]: string[] };
    usersToManager: { [i: string]: string | undefined };
    manageToUsers: { [i: string]: string[] };
    usersToManagerName: { [i: string]: string | undefined };
}

let data: Data = {
    jobs: [],
    users: [],
    groups: [],
    usersToGroups: {},
    usersToManager: {},
    manageToUsers: {},
    usersToManagerName: {},
};

let indexUsers: { [i: string]: MicrosoftGraph.User } = {};


process.on('unhandledRejection', (reason, p) => {
    console.error('Unhandled Rejection at: Promise', p, 'reason:', reason);
    // application specific logging, throwing an error, or other logic here
});



const firedIcons = ['dumpster_fire', 'thisisfine', 'thanos', 'thanos_snap', 'gfto', 'rip', 'notlikethis', 'nothing_to_do_here', 'happening', 'gottarun', 'feelsbadman', 'evilburns', 'downvote', 'disappear', 'cerealspit', 'badtime', 'angry_pepe'];
const hiredIcons = ['very_nice', 'heygirl', 'haveaseat', 'goodnews', 'fonzie', 'excellent', 'catchemall'];
const changeIcons = ['business_cat', 'hr_vest', 'zoidberg', 'notbad', 'indeed', 'feels-good', 'derp', 'clarence', 'caruso', 'bender'];
const jobsIcons = ['theytookourjobs', 'jobs'];

function sendFiredBot(text: string) {
    console.log('sendFiredBot', text);
    // return Promise.resolve(true);
    const icon_emoji = _.sample(firedIcons);
    return slack.send({text, username: 'firedBot', icon_emoji: `:${icon_emoji}:`})
}

function sendHiredBot(text: string) {
    console.log('sendHiredBot', text);
    // return Promise.resolve(true);
    const icon_emoji = _.sample(hiredIcons);
    return slack.send({text, username: 'hiredBot', icon_emoji: `:${icon_emoji}:`})
}

function sendChangeBot(text: string) {
    console.log('sendChangeBot', text);
    // return Promise.resolve(true);
    const icon_emoji = _.sample(changeIcons);
    return slack.send({text, username: 'changeBot', icon_emoji: `:${icon_emoji}:`})
}

function sendJobsBot(text: string) {
    console.log('sendJobsBot', text);
    // return Promise.resolve(true);
    const icon_emoji = _.sample(jobsIcons);
    return slack.send({text, username: 'jobsBot', icon_emoji: `:${icon_emoji}:`})
}

// sendFiredBot('test');
// sendHiredBot('test');
// sendChangeBot('test');


async function sendToTeams(card) {
    // return Promise.resolve(true);
    return request.post(teamsurl, {
        json: card,
    }).then((parsedBody) => {
        console.log(parsedBody);
        return parsedBody;
    });
}

async function writeData() {
    // return Promise.resolve();
    return fs.writeJson('./data.json', data);
}

async function getMore<T>(url: string): Promise<T[]> {
    return request.get(url, {
        headers: {
            'Authorization': `Bearer ${token}`
        },
        json: true
    }).then(async x => {
        // console.log(x);
        if (x['@odata.nextLink']) {
            const more = await getMore(x['@odata.nextLink']);

            x.value.push(...more);

        }

        if (x.value)
            return x.value;

        return x


    })
}

async function get<T>(endpoint: string): Promise<T[]> {
    return getMore<T>(`https://graph.microsoft.com/v1.0/${endpoint}`);
}

async function getSingle<T>(endpoint: string): Promise<T> {
    return getMore<T>(`https://graph.microsoft.com/v1.0/${endpoint}`) as any as T;
}

async function getTeamsOf(userId: string): Promise<(MicrosoftGraph.Group & { '@odata.type': string })[]> {

    return get(`users/${userId}/memberOf`) as any
}

async function doThings(): Promise<boolean> {

    // uncomment if you get the auto thing to work
    // token = await getToken();

    data = await fs.readJson(file);

    await get('groups').then(async x => {
        const groups = (x as any as MicrosoftGraph.Group[]);
        console.log('groups: ', groups.length);

        const newIds = _.differenceBy(groups, data.groups, 'id');
        const notHere = _.differenceBy(data.groups, groups, 'id');

        if (notHere.length) {
            console.log('\n\nnot Here:');
            const nothereDisp = notHere.map(x => x.displayName).join("\n\n");
            console.log(nothereDisp);
            await sendFiredBot(`Not Here Groups:\n${nothereDisp}`);
            await sendToTeams({summary: 'Not Here Groups', title: 'Not Here Groups', text: nothereDisp});

        }

        if (newIds.length) {

            console.log('\n\nnew Groups: ');
            const newGroups = newIds.map(x => x.displayName).join("\n\n");
            console.log(newGroups);
            await sendHiredBot(`New Groups (${newIds.length}):\n${newGroups}`);
            await sendToTeams({summary: 'New Groups', title: 'New Groups', text: newGroups});
            console.log('\n\n');
        }

        const sames = _.intersectionBy(groups, data.groups, 'id');

        const sections = [];
        let samesMsg = '';
        sames.forEach(g => {

            const oldG = data.groups.find(gg => gg.id === g.id);
            if (!oldG) return;

            const diffed = diff(oldG, g);
            if (diffed) {
                const filtered = diffed.filter(d => {
                    return !['onPremisesDomainName','onPremisesSamAccountName','onPremisesNetBiosName','proxyAddresses', 'onPremisesLastSyncDateTime', 'resourceProvisioningOptions', 'isAssignableToRole','securityIdentifier'].includes(d.path[0]);
                });
                if (filtered.length > 0) {
                    console.log(g.displayName, filtered);
                    let msg = `*${g.displayName}* `;
                    const section: { [i: string]: any } = {};
                    section.activityTitle = `*${g.displayName}*`;
                    section.activitySubtitle = 'Changes';

                    let changed = false;

                    filtered.filter(x => x.kind === 'E').forEach(x => {
                        msg += `${x.path.join('.')} from:${x.lhs} to:${x.rhs}\n\n`;
                        changed = true;
                    });
                    filtered.filter(x => x.kind === 'N').forEach(x => {
                        if (x.rhs){
                            changed = true;
                            msg += `New: ${x.path.join('.')} ${x.rhs}\n\n`;
                        }
                    });
                    filtered.filter(x => x.kind === 'D').forEach(x => {
                        msg += `Deleted: ${x.path.join('.')} ${x.rhs}\n\n`;
                        changed = true;

                    });
                    if (changed) {
                        samesMsg += msg;
                        section.text = msg;
                        sections.push(section);
                    }

                }
            }

        });
        if (samesMsg) {
            await sendChangeBot('Group Changes:\n\n' + samesMsg);
            await sendToTeams({summary: 'Group Changes', title: 'Group Changes', sections});
        }


        data.groups = groups;

        return groups;

    });


    const indexedTeams = _.keyBy(data.groups, 'id');

    const usersNew: {
        [k: string]: {
            managerError?: boolean;
            allTeamsGone: boolean;
            userData: MicrosoftGraph.User;
            changes?: any;
            isNew?: MicrosoftGraph.User;
            prevManager?: MicrosoftGraph.User | null;
            manager?: MicrosoftGraph.User | null;
            teams?: any;
            numTeams?: any;
            newTeams?: any;
            missingTeams?: any;
        };
    } = {};

    const users = await get('users').then(async (x) => {
        const users = (x as any as MicrosoftGraph.User[]);
        // console.log(users.filter(u => u.givenName && u.givenName.includes('Kyle')));


        console.log('users:', users.length);

        const newIds = _.differenceBy(users, data.users, 'id');
        const notHere = _.differenceBy(data.users, users, 'id');
        console.log('not Here:', notHere);

        if (notHere.length) {
            const header = `Not Here Users (${notHere.length})`;
            const notHereMsg = notHere.map(u => `*${u.displayName}* ${u.jobTitle} ${u.officeLocation}`).join("\n\n");
            const message = `${header}:\n${notHereMsg}`;
            console.log(message);
            await sendFiredBot(message);
            await sendToTeams({
                summary: `Not Here Users (${notHere.length})`,
                title: `Not Here Users (${notHere.length})`,
                text: notHereMsg
            });
        }

        console.log('new users: ', newIds);
        if (newIds.length) {
            const notHereMsg = `New Here Users (${newIds.length}):\n` + newIds.map(u => `*${u.displayName}* ${u.jobTitle} ${u.officeLocation}`).join("\n");
            console.log(notHereMsg);
            // will be sent later
            // slack.send(notHereMsg);
        }


        const sames = _.intersectionBy(users, data.users, 'id');

        const changes: { [i: string]: deepDiff.IDiff[] } = {};
        sames.forEach(g => {

            const oldG = data.users.find(gg => gg.id === g.id);
            if (!oldG) return;

            const diffed: deepDiff.IDiff[] = diff(oldG, g);
            if (diffed) {
                changes[g.id as string] = diffed;
                console.log(g.displayName, diffed.filter(d => d.path[0] !== 'onPremisesLastSyncDateTime'));
            }

        });


        data.users = users;


        return {users, newIds, changes};


    }).catch(e => {
        // console.trace('catch error');
        console.error(e);
    });


    if (!users) {
        console.error('users if not set');
        return false;
    }

    const changed = users.changes;
    const newIds = _.keyBy(users.newIds, 'id');


    // let i = users.length;
    let c = 0;
    const totalUsers = users.users.length;
    await Bluebird.map(users.users, async (u: MicrosoftGraph.User) => {
        c++;
        if (!u.id) return;
        // if (u.displayName && u.displayName.includes("Tarver")) debugger;
        console.log(`${c}/${totalUsers} ${u.displayName} ${u.id}`);
        const userId: string = u.id;
        usersNew[userId] = {allTeamsGone: false, userData: u, changes: changed[userId], isNew: newIds[userId]};

        await getTeamsOf(u.id).then(teams => {
            // console.log('getting all users teams');
            // console.log(i--);
            teams = teams.filter(nt => nt['@odata.type'] !== '#microsoft.graph.directoryRole');
            // if (u.displayName && u.displayName.includes("Tarver")) teams = [];
            usersNew[userId].teams = teams;

            const teamIds: string[] = teams.map(t => t.id).filter(t => t) as string[];
            // console.log(`${u.displayName} : ${teams.map(t => t.displayName).join(' | ')}`);

            usersNew[userId].numTeams = teamIds.length;

            const userTeams = _.map(data.usersToGroups[u.id as string], id => indexedTeams[id]);

            let newTeams = _.differenceBy(teams, userTeams, 'id').filter(nt => nt['@odata.type'] !== '#microsoft.graph.directoryRole');
            const missingTeams = _.differenceBy(userTeams, teams, 'id').filter(x => x);
            usersNew[userId].newTeams = newTeams;
            usersNew[userId].missingTeams = missingTeams;
            usersNew[userId].allTeamsGone = (userTeams.length > 0 && missingTeams.length === userTeams.length);


            if (newTeams.length) {
                // console.log(teams);
                // console.log(userTeams);
                // console.log(data.usersToGroups[u.id as string]);
                // console.log(u.displayName, ' Added To Teams:', newTeams.map(t => t.displayName));
            }
            if (missingTeams.length) {
                // console.log(teams);
                // console.log(userTeams);
                // console.log(data.usersToGroups[u.id as string]);
                // console.log(u.displayName, ` # teams ${teamIds.length} left, Removed From Teams:`, missingTeams.map(t => t.displayName));
            }

            data.usersToGroups[u.id as string] = teamIds;

            return true;

        }).catch(e => {
            // console.trace('catch error');
            console.error(e);
        });


        indexUsers = _.keyBy(data.users, 'id');

        // console.log('getting all users managers');
        await getSingle(`users/${userId}/manager`)
            .catch((e: StatusCodeError) => {
                if (e.statusCode && e.statusCode === 404) {
                    return undefined;
                }
                console.error('manager error', e);

                throw e;
            })
            .then((manager: MicrosoftGraph.User | undefined) => {

                const savedManagerId: string | undefined = data.usersToManager[userId];
                if (!savedManagerId) {

                    if (manager) {
                        console.log(u.displayName, ' no previous manager, changed to ', manager.displayName);
                        usersNew[userId].prevManager = null;
                        usersNew[userId].manager = manager;
                    } else {
                        // console.log(u.displayName, ' no previous manager, no new manager');
                    }

                } else if (!manager) {

                    if (savedManagerId) {
                        const savedManager: MicrosoftGraph.User = indexUsers[savedManagerId];
                        if (savedManager) {
                            console.log(u.displayName, 'manager changed from ', savedManager.displayName, ' to nothing');
                        } else {
                            console.log(u.displayName, 'manager changed from unknown to nothing');
                        }
                        usersNew[userId].prevManager = savedManager;
                        usersNew[userId].manager = null;
                    } else {
                        // console.log(u.displayName, ' no previous manager, no new manager');
                    }
                } else {

                    const savedManager: MicrosoftGraph.User = indexUsers[savedManagerId];
                    if (!savedManagerId && !manager) {
                        console.log(u.displayName, ' no manager both');
                        usersNew[userId].prevManager = null;
                        usersNew[userId].manager = null;
                    } else if (!manager && savedManager) {
                        console.log(u.displayName, 'manager changed from ', savedManager.displayName, ' to nothing');
                        usersNew[userId].prevManager = savedManager;
                        usersNew[userId].manager = null;
                    } else if (!savedManager) {
                        console.log(u.displayName, '  manager changed to ', manager.displayName);
                        usersNew[userId].prevManager = null;
                        usersNew[userId].manager = manager;
                    } else if (savedManager.id !== manager.id) {
                        console.log(u.displayName, '  manager changed from ', savedManager.displayName, ' to ', manager.displayName);
                        usersNew[userId].prevManager = savedManager;
                        usersNew[userId].manager = manager;
                    } else if (savedManager.id === manager.id) {
                        // same manager
                        // do nothing
                    } else {
                        console.log('Unknown manager condition', savedManagerId, savedManager, manager);
                    }

                }


                data.usersToManager[userId] = manager ? manager.id : undefined;
                // if(!data.usersToManagerName) data.usersToManagerName = {};
                // if(!data.manageToUsers) data.manageToUsers = {};
                data.usersToManagerName[u.displayName] = manager ? manager.id : undefined;
                if (manager) {
                    if (!data.manageToUsers[manager.id]) {
                        data.manageToUsers[manager.id] = [];
                    }
                    data.manageToUsers[manager.id].push(u.displayName);
                }


                return true;


            }).catch(e => {
                console.error(e);
                usersNew[userId].managerError = true;
                // console.trace('catch error');
                notifier.notify('There was an error getting teams');
            })


    }, {concurrency: 5})
        .catch((e: any) => {
            notifier.notify('There was an error getting teams');
            return console.error(e);
        })
        .then(() => {
            console.log('done');
            return true;
        }).catch((e: any) => {
            notifier.notify('There was an error getting teams');
            return console.error(e);
        });

    console.log(`processing all users - ${Object.keys(usersNew).length}`);
    let i = 0;
    const newUsers: string[] = [];
    const changes: string[] = [];
    const fired: string[] = [];

    _.orderBy(usersNew, u => u.userData.displayName)
        .forEach((u) => {
            console.log(`${++i}/${Object.keys(usersNew).length} - ${u.userData.displayName}`);
            const msg: string[] = [];
            let likelyFired = false;
            if (u.allTeamsGone) {
                likelyFired = true;
                console.log('is not a member of any teams anymore');
                msg.push('is not a member of any teams anymore');
            }
            if (u.isNew) {
                msg.push('is new');
            }
            if (u.changes) {
                msg.push('has changes');


                const filtered: deepDiff.IDiff[] = (u.changes as deepDiff.IDiff[])
                    .filter(d => !['proxyAddresses', 'onPremisesLastSyncDateTime', 'resourceProvisioningOptions'].includes(d.path[0]));
                if (filtered.length > 0) {
                    filtered.filter(x => x.kind === 'E').forEach(x => {
                        msg.push(`${x.path.join('.')} from:${x.lhs} to:${x.rhs}`);
                    });
                    filtered.filter(x => x.kind === 'N').forEach(x => {
                        msg.push(`New: ${x.path.join('.')} ${x.lhs}`);
                    });
                    filtered.filter(x => x.kind === 'D').forEach(x => {
                        msg.push(`Deleted: ${x.path.join('.')} ${x.rhs}`);
                    });

                }
            }

            if (!u.managerError) {
                if (u.prevManager !== u.manager) {
                    if (!u.prevManager && !u.manager) {
                        msg.push('has no prev or current manager');
                    } else if (u.prevManager && !u.manager) {
                        likelyFired = true;
                        msg.push(`manager changed from ${u.prevManager.displayName} to nothing`);
                    } else if (!u.prevManager && u.manager) {
                        msg.push(`has no prev and new manager is ${u.manager.displayName}`);
                    } else if (u.prevManager && u.manager) {
                        msg.push(`manager changed from ${u.prevManager.displayName} and new manager is ${u.manager.displayName}`);
                    } else {
                        msg.push('has manager data problems');
                    }
                }
            }

            if (msg.length) {
                if (u.isNew) {
                    newUsers.push(`*${u.userData.displayName}* - ${u.userData.jobTitle} - ${u.userData.officeLocation} - ${msg.join(' | ')}`);
                } else if (likelyFired) {
                    fired.push(`*${u.userData.displayName}* - ${u.userData.jobTitle} - ${u.userData.officeLocation} - ${msg.join(' | ')}`);
                } else {
                    changes.push(`*${u.userData.displayName}* - ${u.userData.jobTitle} - ${u.userData.officeLocation} - ${msg.join(' | ')}`);
                }
            }
        });
    if (newUsers.length > 0) {
        console.log('new users: ', newUsers.join('\n'));
        await sendHiredBot(newUsers.join('\n'));
        await sendToTeams({summary: `New Users`, title: `New Users`, text: newUsers.join('\n\n')});
    }
    if (changes.length > 0) {
        console.log('changes: ', changes.join('\n'));
        await sendChangeBot(changes.join('\n'));
        await sendToTeams({summary: `User Changes`, title: `User Changes`, text: changes.join('\n\n')});
    }
    if (fired.length > 0) {
        console.log('fired: ', fired.join('\n'));
        await sendFiredBot(fired.join('\n'));
        await sendToTeams({summary: `Gone Users`, title: `Gone Users`, text: fired.join('\n\n')});
    }
    if (newUsers.length === 0 && changes.length === 0 && fired.length === 0) {
        console.log('nothing changed');
    }

    const sorter = function <T>(arr: { [x: string]: T }, fn: (x: T) => string) {
        return _.orderBy(_.toPairs(_.countBy(arr, fn)), i => i[1]).reverse();
    };
    const jobTitleStats = sorter(usersNew, (u) => u.userData.jobTitle || 'NO JOB TITLE').map(x => x.join(' - '));
    console.log(jobTitleStats);
    const officeLocationStats = sorter(usersNew, u => u.userData.officeLocation || 'NO OFFICE LOCATION').map(x => x.join(' - '));
    console.log(officeLocationStats);

    const managerToUsers: { [i: string]: number } = {};
    _.forEach(data.usersToManager, (m, k) => {
        if (!indexUsers[k]) {
            delete data.usersToManager[k];
            return;
        }
        if (!managerToUsers[m || 'NO MANAGER']) managerToUsers[m || 'NO MANAGER'] = 0;
        managerToUsers[m || 'NO MANAGER']++;
    });

    const managerStats = _.orderBy(_.toPairs(managerToUsers), x => x[1]).reverse().map(x => {
        const user = indexUsers[x[0]];
        if (!user) {
            return x
        }
        return [user.displayName, x[1]];
    }).map(x => x.join(' - '));
    console.log(managerStats);


    console.log('getting jobs');
    const jobs = await getJobs() as Job[];
    if (!data.jobs) data.jobs = [];

    const newJobs = _.differenceBy(jobs, data.jobs, 'id');
    const deletedJobs = _.differenceBy(data.jobs, jobs, 'id');
    const sameJobs = _.intersectionBy(jobs, data.jobs, 'id');

    if (newJobs.length > 0 && data.jobs.length > 0) {
        const newJobsMessage = newJobs.map(j => `*${j.title}* - ${j.division} - ${j.location} - #${j.id}`).join("\n\n");
        await sendJobsBot(`New Open Jobs (${newJobs.length}/${jobs.length}):\n\n${newJobsMessage}`);
        await sendToTeams({
            summary: `New Open Jobs (${newJobs.length}/${jobs.length})`,
            title: `New Open Jobs (${newJobs.length}/${jobs.length})`,
            text: newJobsMessage
        });
    }

    if (deletedJobs.length > 0 && data.jobs.length > 0) {
        const newJobsMessage = deletedJobs.map(j => `*${j.title}* - ${j.division} - ${j.location} - #${j.id}`).join("\n\n");
        await sendJobsBot(`Deleted Jobs (${deletedJobs.length}/${jobs.length}):\n\n${newJobsMessage}`);
        await sendToTeams({
            summary: `Deleted Open Jobs (${deletedJobs.length}/${jobs.length})`,
            title: `Deleted Open Jobs (${deletedJobs.length}/${jobs.length})`,
            text: newJobsMessage,
        });
    }



    const sections = [];
    let samesMsg = '';
    sameJobs.forEach(g => {

        const oldJob = data.jobs.find(gg => gg.id === g.id);
        if (!oldJob) return;

        const filtered = diff(oldJob, g);
        if (filtered && filtered.length > 0) {
            console.log(g.title, filtered);
            let msg = `*${g.title}* `;
            const section: { [i: string]: any } = {};
            section.activityTitle = g.title;
            section.activitySubtitle = 'Changes';


            filtered.filter(x => x.kind === 'E').forEach(x => {
                msg += `${x.path.join('.')} from:${x.lhs} to:${x.rhs}\n\n`;
            });
            filtered.filter(x => x.kind === 'N').forEach(x => {
                msg += `New: ${x.path.join('.')} ${x.lhs}\n\n`;
            });
            filtered.filter(x => x.kind === 'D').forEach(x => {
                msg += `Deleted: ${x.path.join('.')} ${x.rhs}\n\n`;

            });
            samesMsg += msg;
            section.text = msg;
            sections.push(section);

        }

    });
    if (samesMsg) {
        await sendJobsBot(`Job Changes (${sameJobs.length}/${jobs.length}):\n\n${samesMsg}`);
        await sendToTeams({summary: `Open Jobs Changes (${sameJobs.length}/${jobs.length})`, title: `Open Job Changes (${sameJobs.length}/${jobs.length})`, sections});
    }



    data.jobs = jobs;

    await writeData().catch(e => {
        // console.trace('catch error');
        console.error(e);
    });

    return true;
}


doThings().catch(e => {
    notifier.notify('There was an error getting teams');
    return console.error(e);
}).then(() => {
    console.log('done');
    return process.exit(0);
});
