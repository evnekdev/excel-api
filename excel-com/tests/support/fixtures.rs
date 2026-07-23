use std::fs;
use std::io;
use std::path::{Path, PathBuf};

#[allow(dead_code)]
#[derive(Clone, Copy, Debug)]
pub enum Fixture {
    BlankXlsx,
    BlankXlsm,
    ChartSource,
    PivotSource,
    LocalQuerySource,
}

impl Fixture {
    pub fn path(self) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests")
            .join("fixtures")
            .join(self.name())
    }

    pub fn copy_for_test(self) -> io::Result<TemporaryFixture> {
        self.verify()?;
        let path = std::env::temp_dir().join(format!(
            "excel-com-fixture-{}-{}",
            std::process::id(),
            self.name()
        ));
        fs::copy(self.path(), &path)?;
        Ok(TemporaryFixture { path })
    }

    pub fn verify(self) -> io::Result<()> {
        let bytes = fs::read(self.path())?;
        if !bytes.starts_with(b"PK\x03\x04") || sha256_hex(&bytes) != self.sha256() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "fixture integrity check failed",
            ));
        }
        for required in [
            b"[Content_Types].xml".as_slice(),
            b"xl/workbook.xml",
            b"xl/worksheets/sheet1.xml",
        ] {
            if !bytes
                .windows(required.len())
                .any(|window| window == required)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "fixture archive structure is incomplete",
                ));
            }
        }
        for forbidden in [
            b"externalLinks".as_slice(),
            b"vbaProject",
            b"printerSettings",
            b"hidden",
        ] {
            if bytes
                .windows(forbidden.len())
                .any(|window| window == forbidden)
            {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "fixture has a disallowed archive part",
                ));
            }
        }
        Ok(())
    }

    fn name(self) -> &'static str {
        match self {
            Self::BlankXlsx => "blank.xlsx",
            Self::BlankXlsm => "blank.xlsm",
            Self::ChartSource => "chart-source.xlsx",
            Self::PivotSource => "pivot-source.xlsx",
            Self::LocalQuerySource => "local-query-source.xlsx",
        }
    }

    fn sha256(self) -> &'static str {
        match self {
            Self::BlankXlsx => "0E635B45583C00955F95A44AAC95BC4CE12B28DF8DC659B388E32604B4957FB8",
            Self::BlankXlsm => "0DEF27553F11F38009ED95619F38BBBB3CDFE8C97BF80F32CABFFB5E747B7820",
            Self::ChartSource => "E03895C5E903B68F9519782BAD9C1B0288A303DF125F722DD21C8A95329C9C04",
            Self::PivotSource => "680CE55DB1919C4E8A682F242E176A43AACC03DFAD77B07A8DABF9BB55B60B6F",
            Self::LocalQuerySource => {
                "37821CB413F9C6A26B48F0923DFDA3B36398E28DCA08B48550D284A14CF6E3E8"
            }
        }
    }
}

pub struct TemporaryFixture {
    path: PathBuf,
}

impl TemporaryFixture {
    pub fn path(&self) -> &Path {
        &self.path
    }
}

impl Drop for TemporaryFixture {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.path);
    }
}

fn sha256_hex(input: &[u8]) -> String {
    let mut state = [
        0x6a09e667_u32,
        0xbb67ae85,
        0x3c6ef372,
        0xa54ff53a,
        0x510e527f,
        0x9b05688c,
        0x1f83d9ab,
        0x5be0cd19,
    ];
    let mut padded = input.to_vec();
    padded.push(0x80);
    while (padded.len() + 8) % 64 != 0 {
        padded.push(0);
    }
    padded.extend_from_slice(&((input.len() as u64).wrapping_mul(8)).to_be_bytes());
    for chunk in padded.chunks_exact(64) {
        let mut words = [0_u32; 64];
        for (index, word) in words[..16].iter_mut().enumerate() {
            *word = u32::from_be_bytes(
                chunk[index * 4..index * 4 + 4]
                    .try_into()
                    .expect("chunk word"),
            );
        }
        for index in 16..64 {
            let s0 = words[index - 15].rotate_right(7)
                ^ words[index - 15].rotate_right(18)
                ^ (words[index - 15] >> 3);
            let s1 = words[index - 2].rotate_right(17)
                ^ words[index - 2].rotate_right(19)
                ^ (words[index - 2] >> 10);
            words[index] = words[index - 16]
                .wrapping_add(s0)
                .wrapping_add(words[index - 7])
                .wrapping_add(s1);
        }
        let mut working = state;
        for (index, constant) in K.iter().enumerate() {
            let s1 = working[4].rotate_right(6)
                ^ working[4].rotate_right(11)
                ^ working[4].rotate_right(25);
            let choice = (working[4] & working[5]) ^ ((!working[4]) & working[6]);
            let temp1 = working[7]
                .wrapping_add(s1)
                .wrapping_add(choice)
                .wrapping_add(*constant)
                .wrapping_add(words[index]);
            let s0 = working[0].rotate_right(2)
                ^ working[0].rotate_right(13)
                ^ working[0].rotate_right(22);
            let majority =
                (working[0] & working[1]) ^ (working[0] & working[2]) ^ (working[1] & working[2]);
            let temp2 = s0.wrapping_add(majority);
            working = [
                temp1.wrapping_add(temp2),
                working[0],
                working[1],
                working[2],
                working[3].wrapping_add(temp1),
                working[4],
                working[5],
                working[6],
            ];
        }
        for (value, result) in state.iter_mut().zip(working) {
            *value = value.wrapping_add(result);
        }
    }
    state.iter().map(|value| format!("{value:08X}")).collect()
}

const K: [u32; 64] = [
    0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
    0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
    0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
    0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
    0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
    0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
    0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
    0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2,
];
