import os,tempfile,unittest
from pathlib import Path
from unittest.mock import patch,Mock
import bootstrap

class Bootstrap(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.root=Path(self.tmp.name)
        (self.root/'requirements.txt').write_text('Flask==3.1.3\n',encoding='utf-8')
        self.env=self.root/'.venv';self.python=bootstrap.env_python(self.env)
        self.python.parent.mkdir(parents=True);self.python.write_text('old interpreter')
        self.data=self.root/'personnel';self.data.mkdir();(self.data/'hr.sqlite').write_bytes(b'preserve personnel')
    def tearDown(self):self.tmp.cleanup()
    def test_installed_dependencies_start_without_pip_even_if_pip_is_broken(self):
        with patch.object(bootstrap,'dependencies_ok',return_value=True),patch.object(bootstrap,'pip_ok') as pip,patch.object(bootstrap.subprocess,'run') as run:
            self.assertEqual(bootstrap.prepare(self.root),self.python);pip.assert_not_called();run.assert_not_called()
    def test_broken_pip_is_rebuilt_once_before_installing(self):
        def build(env):
            py=bootstrap.env_python(env);py.parent.mkdir(parents=True);py.write_text('new interpreter')
        with patch.object(bootstrap,'dependencies_ok',side_effect=[False,True]),patch.object(bootstrap,'pip_ok',side_effect=[False,True]),patch.object(bootstrap.venv.EnvBuilder,'create',side_effect=build),patch.object(bootstrap.subprocess,'run',return_value=Mock(returncode=0)) as run:
            bootstrap.prepare(self.root)
            self.assertEqual(self.python.read_text(),'new interpreter')
            backups=list(self.root.glob('.venv-backup-*'));self.assertEqual(len(backups),1)
            self.assertEqual(bootstrap.env_python(backups[0]).read_text(),'old interpreter')
            args=run.call_args.args[0];self.assertIn('pip',args);self.assertNotIn('--force-reinstall',args)
        self.assertEqual((self.data/'hr.sqlite').read_bytes(),b'preserve personnel')
    def test_failed_rebuild_restores_old_environment(self):
        def fail(env):
            env.mkdir();(env/'partial').write_text('failed attempt');raise RuntimeError('ensurepip failed')
        with patch.object(bootstrap.venv.EnvBuilder,'create',side_effect=fail):
            with self.assertRaisesRegex(RuntimeError,'ensurepip failed'):bootstrap.rebuild_environment(self.env)
        self.assertEqual(self.python.read_text(),'old interpreter');self.assertEqual(len(list(self.root.glob('.venv-failed-*'))),1)
    def test_dependency_install_failure_is_reported_without_retry_loop(self):
        with patch.object(bootstrap,'dependencies_ok',return_value=False),patch.object(bootstrap,'pip_ok',return_value=True),patch.object(bootstrap.subprocess,'run',return_value=Mock(returncode=2)) as run:
            with self.assertRaisesRegex(bootstrap.SetupError,'pip ma loi 2'):bootstrap.prepare(self.root)
            run.assert_called_once();self.assertIn('--force-reinstall',run.call_args.args[0])
    def test_data_directory_inside_environment_blocks_replacement(self):
        with patch.dict(os.environ,{'EMGA_DATA_DIR':str(self.env/'data')}):
            with self.assertRaisesRegex(bootstrap.SetupError,'EMGA_DATA_DIR'):bootstrap.rebuild_environment(self.env)
        self.assertEqual(self.python.read_text(),'old interpreter')
    def test_force_repair_rebuilds_even_when_dependencies_present(self):
        with patch.object(bootstrap,'rebuild_environment') as rebuild,patch.object(bootstrap,'dependencies_ok',return_value=True),patch.object(bootstrap.subprocess,'run',return_value=Mock(returncode=0)):
            bootstrap.prepare(self.root,repair=True);rebuild.assert_called_once_with(self.env)
    def test_probe_checks_real_versions_and_imports(self):
        import sys
        self.assertTrue(bootstrap.dependencies_ok(Path(sys.executable),[('Flask','3.1.3','flask')]))
        self.assertFalse(bootstrap.dependencies_ok(Path(sys.executable),[('Flask','0.0.0','flask')]))

if __name__=='__main__':unittest.main()
