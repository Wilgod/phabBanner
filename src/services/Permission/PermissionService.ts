import { IPermissionService } from "./IPermissionService";
import { RoleTypeKind } from "@pnp/sp/security/types";
import { IWeb } from "@pnp/sp/webs";
import { IList } from "@pnp/sp/lists";
import { error } from "jquery";

export class PermissionService implements IPermissionService {
    constructor(private web: IWeb) {
    }

    public async setFolderPermission(listID: string, users: any[], accessRight: string): Promise<boolean> {
        let result: boolean = false;

        const list: IList = await this.web.lists.getById(listID);
        if (users && users.length > 0) {
            await list.breakRoleInheritance(false, false);
            switch (accessRight) {
                case "edit":
                    return await this.setListPermission(this.web, list, 6, users);
                    break;
                case "read":
                    return await this.setListPermission(this.web, list, 2, users);
                    break;
                default:
                    break;
            }
        }

        return result
    }

    public async setResetPermission(listID: string): Promise<boolean> {
        let result: boolean = false;

        const list: IList = await this.web.lists.getById(listID);
        await list.resetRoleInheritance()
        
        return result
    }

    public async setUserInGroup(listID: string, groupId: number, users: any[]): Promise<boolean> {
        let result: boolean = false;

        const list: IList = await this.web.lists.getById(listID);
        if (users && users.length > 0) {
            await list.breakRoleInheritance(true, true);
            const promises = [];
            for (let i = 0; i <= users.length - 1; i++) {
                promises.push(
                this.web.siteGroups
                .getById(groupId)
                .users.add(users[i].loginName))
            }

            await Promise.all(promises).then((siteUsers) => {
                result = true;
            }).catch((err) => {
                throw err;
            })
        }

        return result;
    }

    private async setListPermission(web: IWeb, list: IList, roleTypeKind: RoleTypeKind, users: any[]): Promise<boolean> {
        let result: boolean = false;

        await web.roleDefinitions.getByType(roleTypeKind)()
            .then(async (rd) => {
                const promises = [];
                for (let i = 0; i <= users.length - 1; i++) {
                    if (!users[i].loginName.includes("membership")) {
                      promises.push(list.roleAssignments.add(users[i].id, rd.Id));
                    } else {
                      await web
                        .ensureUser(users[i].loginName)
                        .then(async (user) => {
                          promises.push(
                            list.roleAssignments.add(user.data.Id, rd.Id)
                          );
                        })
                        .catch((err) => console.log("ensure user err", err));
                    }
                }

                await Promise.all(promises).then(() => {
                    result = true;
                }).catch((err) => {throw err});
            })

        return result;
    }
}
