/*::
declare module 'exit-on-epipe' {};

declare module 'harb' { declare var exports:HARBModule;  };
declare module '../' { declare var exports:HARBModule; };

declare module 'commander' { declare var exports:any; };

type SSFDate = {
	D:number; T:number;
	y:number; m:number; d:number; q:number;
	H:number; M:number; S:number; u:number;
};

type SSFTable = {[key:number|string]:string};
declare module 'ssf' {
	declare function parse_date_code(v:number,opts:any):?SSFDate;
	declare var _table:SSFTable;
	declare function format(fmt:any, value:any):string;
};

type BPResult = any;
declare module 'babyparse' {
	declare function parse(str:string, opts:any):BPResult;
};

type Data = string | Array<number> | Buffer;
type CPIndex = number|string;
declare module './dist/cpharb' {
	declare var utils:{
		decode(cp:CPIndex, data:Data):string;
	};
};

declare module 'xlsx' {
	declare var exports: any;
};
*/
