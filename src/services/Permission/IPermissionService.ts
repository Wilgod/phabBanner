
export interface IPermissionService {
    setFolderPermission: (listID: string, users: any[], accessRight: string) => Promise<boolean>;
    setResetPermission: (listID: string) => Promise<boolean>;
    setUserInGroup:(listID: string, groupId: number, users: any[]) => Promise<boolean>;
}