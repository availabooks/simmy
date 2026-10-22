/**
 * Package the Claude skill (skill/simmy/) for the docs site: <out>/skill/simmy.zip, which
 * holds a simmy/ folder ready to unzip into a skills folder or upload, and the loose files
 * under <out>/skill/simmy/ for anyone who wants to read them first.
 */
import { cpSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { crc32, deflateRawSync } from "node:zlib";

const SKILL_DIR = "skill/simmy";

/** A zip archive of these files, deflated. Enough of the format for any unzip tool. */
export function zip(files: { name: string; data: Buffer }[]): Buffer {
  const locals: Buffer[] = [];
  const centrals: Buffer[] = [];
  let offset = 0;
  for (const { name, data } of files) {
    const path = Buffer.from(name, "utf8");
    const packed = deflateRawSync(data);
    const crc = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4); // version needed
    local.writeUInt16LE(0x0800, 6); // UTF-8 names
    local.writeUInt16LE(8, 8); // deflate
    local.writeUInt32LE(0, 10); // time and date: 1980, so the zip is the same on every build
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(packed.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(path.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(packed.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(path.length, 28);
    central.writeUInt32LE(0o100644 * 0x10000, 38); // a regular file, rw-r--r--
    central.writeUInt32LE(offset, 42);
    locals.push(local, path, packed);
    centrals.push(central, path);
    offset += local.length + path.length + packed.length;
  }
  const directory = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, directory, end]);
}

export function writeSkill(outDir: string): void {
  const names = readdirSync(SKILL_DIR).sort();
  const out = join(outDir, "skill");
  mkdirSync(out, { recursive: true });
  cpSync(SKILL_DIR, join(out, "simmy"), { recursive: true });
  const files = names.map((name) => ({ name: `simmy/${name}`, data: readFileSync(join(SKILL_DIR, name)) }));
  writeFileSync(join(out, "simmy.zip"), zip(files));
}
